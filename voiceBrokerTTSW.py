#This is a version im working on just usng tts-wrapper. Im not sure how event emitting works or the correct calls to speak etc

import pythoncom
import json
import logging
import winreg
import json
import argparse
from tts_wrapper import PollyTTS, PollyClient, MicrosoftTTS, MicrosoftClient, WatsonTTS, WatsonClient, GoogleTTS, GoogleClient, ElevenLabsTTS, ElevenLabsClient


class VoiceBroker:
    _public_methods_ = ['Speak', 'Pause', 'Resume', 'GetVoices', 'SetVoice', 'SetInterest', 'WaitForNotifyEvent']
    _reg_progid_ = "VoiceBroker.Engine"
    _reg_clsid_ = "{4DFFD59B-4DF3-4366-B053-DFF9BE002EFB}" 


    def __init__(self, settings_file, credentials_file):
        self.settings = self.load_settings(settings_file)
        self.credentials = self.load_credentials(credentials_file)
        self.tts_engines = {}
        self.pOutputSite = None
        self.event_interests = {}
        self.register_tts_engines()

    def load_settings(self, file_path):
        with open(file_path, 'r') as file:
            return json.load(file)

    def load_credentials(self, file_path):
        with open(file_path, 'r') as file:
            return json.load(file)

    def register_tts_engines(self):
        required_engines = set(self.settings["Voices"].keys())
        for engine in required_engines:
            if engine == "polly":
                creds = self.credentials.get('polly', {})
                client = PollyClient(credentials=(creds.get('aws_key_id'), creds.get('aws_access_key')))
                self.tts_engines[engine] = PollyTTS(client=client)
            elif engine == "microsoft":
                creds = self.credentials.get('microsoft', {})
                client = MicrosoftClient(credentials=creds.get('token'), region=creds.get('region'))
                self.tts_engines[engine] = MicrosoftTTS(client=client)
            elif engine == "elevenlabs":
                creds = self.credentials.get('elevenlabs', {})
                client = ElevenLabsClient(credentials=creds.get('api_key'))
                self.tts_engines[engine] = ElevenLabsTTS(client=client)
            elif engine == "watson":
                creds = self.credentials.get('watson', {})
                client = WatsonClient(credentials=(creds.get('api_key'), creds.get('region'), creds.get('instance_id')))
                self.tts_engines[engine] = WatsonTTS(client=client)
            elif engine == "google":
                creds = self.credentials.get('google', {})
                client = GoogleClient(credentials=creds.get('creds_path'))
                self.tts_engines[engine] = GoogleTTS(client=client)

    def _extract_text_from_fragments(self, pTextFragList):
        # This is a simplified example. Actual implementation depends on the structure of pTextFragList.
        text_fragments = []
        fragment = pTextFragList
        while fragment:
            text_fragments.append(fragment.Text)
            fragment = fragment.pNext
        return " ".join(text_fragments)

    def Speak(self, dwSpeakFlags, rguidFormatId, pWaveFormatEx, pTextFragList, pOutputSite):
        self.pOutputSite = pOutputSite

        # Extract text from the pTextFragList
        text = self._extract_text_from_fragments(pTextFragList)

        # For simplicity, we assume the first engine in the list is used
        engine_name, engine = next(iter(self.tts_engines.items()))

        def handle_event(word, start_time):
            self.EventNotify(1, start_time, word)

        ssml_text = engine.ssml.add(text)
        engine.start_playback_with_callbacks(ssml_text, callback=handle_event)

    def EventNotify(self, stream_number, start_time, text):
        # Emit events via the pOutputSite
        if self.pOutputSite:
            self.pOutputSite.OnWord(stream_number, start_time, len(text))

    def SetInterest(self, event_id, enabled):
        self.event_interests[event_id] = enabled
        
    def WaitForNotifyEvent(self, timeout):
        pass  # Placeholder, implement as needed
    
    def GetOutputFormat(self, pTargetFmtId, pTargetWaveFormatEx, ppCoMemActualFmtId, ppCoMemActualWaveFormatEx):
        # Implement the GetOutputFormat method here
        pass


    def Pause(self):
        for engine in self.tts_engines.values():
            engine.pause_audio()

    def Resume(self):
        for engine in self.tts_engines.values():
            engine.resume_audio()

    def GetVoices(self):
        voices = []
        for engine in self.tts_engines.values():
            voices.extend(engine.get_voices())
        return voices

    def SetVoice(self, voice_id):
        for engine in self.tts_engines.values():
            engine.set_voice(voice_id)


def register_voice(voiceid, name, language, gender, vendor):
    # Paths for 64-bit and 32-bit applications
    paths = [
        f"SOFTWARE\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}",
        f"SOFTWARE\\WOW6432Node\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}"
    ]

    for path in paths:
        # Create the registry key for the voice
        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path)

        # Set the attributes of the voice
        winreg.SetValue(key, "", winreg.REG_SZ, name)
        winreg.SetValue(key, "409", winreg.REG_SZ, language)
        winreg.SetValue(key, "CLSID", winreg.REG_SZ, "{4DFFD59B-4DF3-4366-B053-DFF9BE002EFB}")

        # Create the registry key for the attributes
        attr_key = winreg.CreateKey(key, "Attributes")
        winreg.SetValue(attr_key, "Language", winreg.REG_SZ, language)
        winreg.SetValue(attr_key, "Gender", winreg.REG_SZ, gender)
        winreg.SetValue(attr_key, "Vendor", winreg.REG_SZ, vendor)
        winreg.SetValue(attr_key, "Name", winreg.REG_SZ, name)

        # Close the registry keys
        winreg.CloseKey(attr_key)
        winreg.CloseKey(key)
        
def register_tts_voices():
    with open('settings.json', 'r') as file:
        settings = json.load(file)

    for engine, voices in settings.items():
        for voice in voices:
            register_voice(**voice)


def unregister_com_server():
    # Unregister the classes using the win32com provided utility
    win32com.server.register.UnregisterClasses(VoiceBroker)
    
    app_name = 'VoiceBroker'
    paths = [
        r"SOFTWARE\Classes\AppID",  # Typical for 32-bit on 32-bit machines or 64-bit on 64-bit machines
        r"SOFTWARE\Wow6432Node\Classes\AppID"  # Typical for 32-bit on 64-bit machines
    ]
    total_deletions = 0
    errors = []

    for path in paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_ALL_ACCESS) as key:
                subkeys = []
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkeys.append(subkey_name)
                        i += 1
                    except WindowsError:
                        break

                for subkey_name in subkeys:
                    try:
                        subkey = winreg.OpenKey(key, subkey_name)
                        try:
                            value, _ = winreg.QueryValueEx(subkey, "AppID")
                            if value.lower() == app_name.lower():
                                winreg.DeleteKey(key, subkey_name)
                                print(f"Deleted {subkey_name} in {path}")
                                total_deletions += 1
                        except FileNotFoundError:
                            pass
                        finally:
                            winreg.CloseKey(subkey)
                    except WindowsError as e:
                        errors.append(f"Failed to delete subkey {subkey_name} in {path}: {str(e)}")

        except WindowsError as e:
            errors.append(f"Failed to open or modify registry at {path}: {str(e)}")

    print(f"Total deletions: {total_deletions}")
    if errors:
        print("Errors encountered:")
        for error in errors:
            print(error)

def unregister_voice(voiceid):
    # Paths for 64-bit and 32-bit applications
    paths = [
        f"SOFTWARE\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}",
        f"SOFTWARE\\WOW6432Node\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}"
    ]

    for path in paths:
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, path)
            print(f"Unregistered voice: {voiceid} from {path}")
        except Exception as e:
            print(f"Failed to unregister voice {voiceid} from {path}: {str(e)}")

def unregister_tts_voices():
    with open('settings.json', 'r') as file:
        settings = json.load(file)

    for engine, voices in settings.items():
        for voice in voices:
            unregister_voice(voice['voiceid'])
            

if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    parser = argparse.ArgumentParser(description="VoiceBroker COM Server")
    parser.add_argument('action', choices=['register', 'unregister'], help="Register or unregister the COM server")
    args = parser.parse_args()

    if args.action == 'register':
        win32com.server.register.UseCommandLine(VoiceBroker)
        register_com_server()
        register_tts_voices()
    elif args.action == 'unregister':
        unregister_com_server()
        unregister_tts_voices()
