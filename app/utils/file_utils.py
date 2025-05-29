# app/utils/file_utils.py

from supabase import Client

def create_signed_url(supabase: Client, bucket: str, path: str, expires_in: int) -> str:
    """
    Gera uma URL assinada para um objeto privado no Storage do Supabase.
    - supabase: instância do Client
    - bucket: nome do bucket (ex: 'relatorios-qag')
    - path: caminho do objeto dentro do bucket
    - expires_in: tempo em segundos para expirar a URL
    """
    resp = supabase.storage.from_(bucket).create_signed_url(path, expires_in)
    if resp.get("error"):
        raise Exception(f"Erro ao criar signed URL: {resp['error']['message']}")
    return resp["signedURL"]
