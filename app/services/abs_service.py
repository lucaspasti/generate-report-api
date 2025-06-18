# abs_service.py

import os
import io
import tempfile
from abc import ABC
from datetime import datetime
from pytz import timezone
from collections import OrderedDict

import numpy as np
import pandas as pd
from supabase import Client
from fastapi import HTTPException
from docxtpl import DocxTemplate, Subdoc, InlineImage
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm


class AbstractService(ABC):
    def __init__(
        self,
        supabase: Client,
        payload,
        TEMPLATE_PATH: str,
        BUCKET: str,
        tipo_relatorio: str,
        nome_formulario: str,
        Response,
        Request,
    ):
        self.supabase = supabase
        self.payload = payload
        self.TEMPLATE_PATH = TEMPLATE_PATH
        self.BUCKET = BUCKET
        self.tipo_relatorio = tipo_relatorio
        self.periodicidade = payload.periodicidade
        self.data_str = payload.data_campanha.isoformat()
        self.tz_br = timezone("America/Sao_Paulo")
        self.now_br = datetime.now(self.tz_br)
        self.object_key = f"{payload.ativo_id}/{self.data_str}/{self.now_br:%Y-%m-%d_%H-%M-%S}.docx"
        self.tmp_dir = tempfile.gettempdir()
        self.server_path = os.path.join(
            self.tmp_dir, os.path.basename(self.object_key))

        self.document = DocxTemplate(self.TEMPLATE_PATH)
        self.ativo, self.configuracoes, self.form = self._get_data(
            nome_formulario)
        resultados = self.form.data[0]["resultados"]
        self.df_resultados = pd.DataFrame(resultados).fillna("Indisponível")

        # parâmetros específicos (podem vir do payload)
        self.parametros = getattr(payload, "parametros", [])
        self.parametros_metais_pesados = getattr(
            payload, "parametros_metais_pesados", [])
        self.solventes = getattr(payload, "solventes", [])

        self.lab = self.form.data[0]
        self.contexto = {}
        self.Response = Response
        self.Request = Request

    def _get_data(self, nome_formulario: str):
        ativo = (
            self.supabase.table("ativos")
            .select("*")
            .eq("id", str(self.payload.ativo_id))
            .execute()
        )
        if not ativo.data:
            raise HTTPException(404, detail="Ativo não encontrado")

        configuracoes = (
            self.supabase.table("configuracao_formulario_ativos")
            .select("*")
            .eq("ativo_id", str(self.payload.ativo_id))
            .eq("tipo_formulario", nome_formulario)
            .execute()
        )
        if not configuracoes.data:
            raise HTTPException(
                404, detail="Configuração do formulário não encontrada")

        form = (
            self.supabase.table(nome_formulario)
            .select("*")
            .eq("ativo_id", str(self.payload.ativo_id))
            .eq("campanha_de_coleta", self.data_str)
            .execute()
        )
        if not form.data:
            raise HTTPException(404, detail="Campanha não encontrada")

        return ativo, configuracoes, form

    def render_document(self):
        """Renderiza e salva o documento no caminho temporário."""
        self.document.render(self.contexto)
        self.document.save(self.server_path)

    def upload_and_register_report(self):
        """Faz upload do DOCX para o Storage e registra no Supabase."""
        with open(self.server_path, "rb") as f:
            data = f.read()

        self.supabase.storage.from_(self.BUCKET).upload(
            self.object_key,
            data,
            {"contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
        )
        public_url = self.supabase.storage.from_(
            self.BUCKET).get_public_url(self.object_key)

        try:
            self.supabase.table("relatorios").insert(
                {
                    "nome_relatorio": self.payload.nome_relatorio,
                    "descricao_relatorio": self.payload.descricao_relatorio,
                    "ativo_id": str(self.payload.ativo_id),
                    "user_id": str(self.payload.user_id),
                    "tipo_relatorio": self.tipo_relatorio,
                    "url_relatorio": public_url,
                }
            ).execute()
        except Exception as e:
            return self.Request(sucesso=False, mensagem=f"Erro ao registrar o relatório: {e}")

        return self.Response(sucesso=True, mensagem="Relatório gerado e registrado com sucesso.")

    def _new_subdoc(self) -> Subdoc:
        """Cria um Subdoc para injeção no contexto."""
        return self.document.new_subdoc()

    def inserir_tabela_agrupada(
        self,
        data: list,
        parametros: list,
        col_group: str = "Ponto",
        col_value: str = "Profundidade",
        style: str = "Table Grid",
        align=WD_TABLE_ALIGNMENT.CENTER,
        ctx_key: str = None,
    ):
        """
        Gera uma tabela agrupada por col_group, com merge vertical nessa coluna.
        data: lista de dicts
        parametros: colunas adicionais no cabeçalho
        ctx_key: chave para salvar o Subdoc em self.contexto
        """
        subdoc = self._new_subdoc()
        n_cols = 2 + len(parametros)
        table = subdoc.add_table(rows=1, cols=n_cols)
        table.style = style
        table.alignment = align

        # cabeçalho
        hdr = table.rows[0].cells
        hdr[0].text = col_group
        hdr[1].text = col_value
        for i, p in enumerate(parametros, start=2):
            hdr[i].text = p
        for cell in hdr:
            for par in cell.paragraphs:
                for run in par.runs:
                    run.font.bold = True

        # agrupa mantendo a ordem de inserção
        grupos = OrderedDict()
        for row in data:
            grupos.setdefault(row[col_group], []).append(row)

        # preenche as linhas e faz merge vertical
        cur = 1
        for chave, rows in grupos.items():
            start = cur
            for r in rows:
                cells = table.add_row().cells
                cells[1].text = str(r[col_value])
                for j, p in enumerate(parametros, start=2):
                    cells[j].text = str(r.get(p, ""))
                cur += 1
            end = cur - 1

            top = table.cell(start, 0)
            for rr in range(start + 1, end + 1):
                top = top.merge(table.cell(rr, 0))
            top.text = str(chave)

        if ctx_key:
            self.contexto[ctx_key] = subdoc
        return subdoc

    def inserir_tabela_media(
        self,
        df: pd.DataFrame,
        parametros: list,
        profundidade_col: str = "Profundidade",
        categorias: list = None,
        style: str = "Table Grid",
        align=WD_TABLE_ALIGNMENT.CENTER,
        ctx_key: str = None,
    ):
        """
        Gera tabela de médias com dois níveis de cabeçalho:
        - linha 0: 'Parâmetro' + colspan categorias
        - linha 1: cada categoria
        Corpo: uma linha por parâmetro.
        """
        categorias = categorias or ["Total", "Superfície", "Meio", "Fundo"]
        subdoc = self._new_subdoc()
        n_rows = 2 + len(parametros)
        n_cols = 1 + len(categorias)
        table = subdoc.add_table(rows=n_rows, cols=n_cols)
        table.style = style
        table.alignment = align

        # cabeçalho nível 0
        hdr0 = table.rows[0].cells
        hdr0[0].text = "Parâmetro"
        span = hdr0[1]
        for i in range(2, n_cols):
            span = span.merge(hdr0[i])
        span.text = "Média"
        for cell in (hdr0[0], span):
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.bold = True

        # cabeçalho nível 1
        hdr1 = table.rows[1].cells
        hdr1[0].text = ""
        for idx, cat in enumerate(categorias, start=1):
            hdr1[idx].text = cat
            for p in hdr1[idx].paragraphs:
                for r in p.runs:
                    r.font.bold = True

        # corpo
        df_copy = df.copy()
        df_copy[profundidade_col] = df_copy[profundidade_col].astype(str)
        for i, param in enumerate(parametros):
            row_cells = table.rows[2 + i].cells
            row_cells[0].text = param
            série = pd.to_numeric(df_copy[param], errors="coerce")
            médias = [
                série.mean(),
                série[df_copy[profundidade_col] == categorias[1]].mean(),
                série[df_copy[profundidade_col] == categorias[2]].mean(),
                série[df_copy[profundidade_col] == categorias[3]].mean(),
            ]
            for col_idx, val in enumerate(médias, start=1):
                row_cells[col_idx].text = "" if pd.isna(val) else f"{val:.2f}"

        if ctx_key:
            self.contexto[ctx_key] = subdoc
        return subdoc

    def _figs_to_inline_images(self, figs, width=Cm(12), height=Cm(6)):
        images = []
        for fig in figs:
            buf = io.BytesIO()
            fig.savefig(buf, format="PNG", dpi=120, bbox_inches="tight")
            buf.seek(0)
            images.append(InlineImage(self.document, buf,
                          width=width, height=height))
        return images

    def _percentual_conformidade(self, df, parametros, vmp_dict):
        conformes = []
        for _, row in df.iterrows():
            cls = row["Classe"]
            limites = vmp_dict.get(cls, {})
            for p in parametros:
                if p not in limites:
                    continue
                try:
                    vmp = float(limites[p])
                    val = float(row.get(p, np.nan))
                    conformes.append(val <= vmp)
                except:
                    pass
        return (sum(conformes) / len(conformes) * 100) if conformes else 0.0
