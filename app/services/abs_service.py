import os
import io
import datetime
from datetime import datetime
import pandas as pd
from pytz import timezone
from supabase import Client
from fastapi import HTTPException
from docxtpl import DocxTemplate
import tempfile
from abc import ABC, abstractmethod


class AbstractService(ABC):

    def __init__(self, supabase: Client, payload, TEMPLATE_PATH, BUCKET, parametros, tipo_relatorio, Response, Request):
        self.supabase = supabase
        self.payload = payload
        self.TEMPLATE_PATH = TEMPLATE_PATH
        self.BUCKET = BUCKET
        self.periodicidade = self.payload.periodicidade
        self.data_str = payload.data_campanha.isoformat()
        self.tz_br = timezone("America/Sao_Paulo")
        self.now_br = datetime.now(self.self.tz_br)
        self.object_key = f"{self.payload.ativo_id}/{self.data_str}/{self.now_br:%Y-%m-%d_%H-%M-%S}.docx"
        self.tmp_dir = tempfile.gettempdir()
        self.server_path = os.path.join(
            self.tmp_dir, os.path.basename(self.object_key))
        self.document = DocxTemplate(self.TEMPLATE_PATH)
        self.ativo, self.configuracoes, self.form = self._get_data()
        self.resultados = self.form.data[0]["resultados"]
        self.df_resultados = pd.DataFrame(
            self.resultados).fillna("Indisponível")
        self.parametros = parametros
        self.lab = self.form.data[0]
        self.contexto = {}
        self.tipo_relatorio = tipo_relatorio
        self.Response = Response
        self.Request = Request

    def _get_data(self, nome_formulario):
        ativo = (
            self.supabase.table("ativos").select(
                "*").eq("id", str(self.payload.ativo_id)).execute()
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

    @abstractmethod
    def render_document(self):
        self.document.render(self.contexto)
        self.document.save(self.server_path)

    @abstractmethod
    def upload_and_register_report(self):
        with open(self.server_path, "rb") as f:
            data = f.read()
        self.supabase.storage.from_(self.BUCKET).upload(
            self.object_key,
            data,
            {
                "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
        )

        # 9) pegar url publica
        public_url = self.supabase.storage.from_(
            self.BUCKET).get_public_url(self.object_key)

        # 10) Insere registro na tabela `relatorios`
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
            return self.Request(
                sucesso=False, mensagem=f"Erro ao registrar o relatório: {str(e)}"
            )

        finally:
            # nenhum erro: devolvemos sucesso
            return self.Response(
                sucesso=True, mensagem="Relatório gerado e registrado com sucesso."
            )
