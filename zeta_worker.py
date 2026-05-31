"""Worker module para multiprocessing - debe estar separado del GUI."""
import os
import re
import unicodedata

BINARY_EXT = ('.png','.jpg','.jpeg','.gif','.bmp','.ico','.dll','.exe','.so','.dylib',
              '.wav','.ogg','.mp3','.flac','.zip','.rar','.7z','.tar','.gz','.bz2',
              '.mp4','.avi','.mkv','.mov','.pdf','.doc','.docx','.xls','.xlsx',
              '.pyc','.pyo','.class','.o','.obj','.bin','.pack','.idx','.db',
              '.sqlite','.lnk','.pdb','.lib','.exp','.woff','.woff2','.ttf','.otf')

MAX_SIZE = 5 * 1024 * 1024  # 5MB

def _normalize(text):
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode().lower()

def search_chunk(args):
    """Función ejecutada en proceso separado. Retorna lista de rutas con coincidencias."""
    paths, keywords_norm, whole_word = args
    results = []
    # Preparar patrones simples normalizados
    for path in paths:
        try:
            size = os.path.getsize(path)
            if size == 0 or size > MAX_SIZE:
                continue
            with open(path, 'rb') as f:
                raw = f.read()
            # Detección binaria rápida
            if b'\x00' in raw[:512]:
                continue
            # Decodificar y normalizar de una vez
            text = _normalize(raw.decode('utf-8', errors='ignore'))
            
            found = True
            for kw in keywords_norm:
                if whole_word:
                    # Buscar palabra completa con regex simple
                    if not re.search(r'(?<![a-z0-9])' + re.escape(kw) + r'(?![a-z0-9])', text):
                        found = False
                        break
                else:
                    if kw not in text:
                        found = False
                        break
            if found:
                # Contar ocurrencias del primer keyword para el score
                score = text.count(keywords_norm[0])
                results.append((score, path))
        except Exception:
            pass
    return results
