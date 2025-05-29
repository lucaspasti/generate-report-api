from supabase import create_client
from app.config import settings

_supabase = None


def get_supabase_client():
    global _supabase
    if not _supabase:
        _supabase = create_client(settings.SUPABASE_URL, settings.SUPABASE_KEY)
    return _supabase
