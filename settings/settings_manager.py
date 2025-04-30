import os
import json

class SettingsManager:
    def __init__(self, settings_file="settings.json"):
        self.settings_file = settings_file
        self.settings = self.load_settings()
        
    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    return json.load(f)
            else:
                return {
                    'file_to_monitor': '',
                    'api_url': '',
                    'api_key': ''
                }
        except Exception:
            return {
                'file_to_monitor': '',
                'api_url': '',
                'api_key': ''
            }
    
    def save_settings(self, settings):
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
            return True
        except Exception:
            return False