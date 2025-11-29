# Translation helper - use this to translate your UI
# 
# Example usage in main.py:
# from translate_ui import apply_translations
# 
# Then in __init__ or where you setup UI:
# apply_translations(self, translator)
#
# This file contains mappings of all UI elements that need translation

def apply_translations(window, t):
    """
    Apply translations to all UI elements
    t is the translator instance with t() method
    """
    pass  # Will be filled by converting existing strings
    
# Translation keys reference:
# t('app.title') -> "PSA-DIAG FREE"
# t('buttons.download') -> "Download" / "Télécharger"
# t('labels.ready') -> "Ready" / "Prêt"
# t('messages.download.title') -> "Download" / "Téléchargement"
# t('messages.download.success', version='9.85') -> "Diagbox 9.85 downloaded successfully"
