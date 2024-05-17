import PySimpleGUI as sg
import json
import logging
from tts_wrapper import PollyClient, PollyTTS, MicrosoftClient, MicrosoftTTS, WatsonClient, WatsonTTS, GoogleClient, GoogleTTS, ElevenLabsClient, ElevenLabsTTS

# Load credentials from file
def load_credentials():
    try:
        with open('credentials.json', 'r') as f:
            credentials = json.load(f)
        logging.debug("Credentials loaded from file.")
        return credentials
    except FileNotFoundError:
        logging.error("Credentials file not found. Ensure 'credentials.json' is in the correct location.")
        return {}

# Load settings from file
def load_settings():
    try:
        with open('settings.json', 'r') as f:
            settings = json.load(f)
        logging.debug("Settings loaded from file.")
        return settings
    except FileNotFoundError:
        logging.error("Settings file not found. Ensure 'settings.json' is in the correct location.")
        return {}

# Create TTS client based on the service type
def create_tts_client(service, credentials):
    if service == "polly":
        creds = credentials.get('polly', {})
        client = PollyClient(credentials=(creds.get('aws_key_id'), creds.get('aws_access_key')))
        tts = PollyTTS(client=client)
    elif service == "microsoft":
        creds = credentials.get('microsoft', {})
        client = MicrosoftClient(credentials=creds.get('token'), region=creds.get('region'))
        tts = MicrosoftTTS(client=client)
    elif service == "watson":
        creds = credentials.get('watson', {})
        client = WatsonClient(credentials=(creds.get('api_key'), creds.get('api_url')))
        tts = WatsonTTS(client=client)
    elif service == "google":
        creds = credentials.get('google', {})
        client = GoogleClient(credentials=creds.get('creds_path'))
        tts = GoogleTTS(client=client)
    elif service == "elevenlabs":
        creds = credentials.get('elevenlabs', {})
        client = ElevenLabsClient(credentials=creds.get('api_key'))
        tts = ElevenLabsTTS(client=client)
    else:
        raise ValueError("Unsupported TTS service")
    return tts

def play_voice(tts, voice_id, lang_code, text="Hello, world!"):
    tts.set_voice(voice_id, lang_code)
    ssml_text = tts.ssml.add(text)
    tts.speak(ssml_text)

def main():
    credentials = load_credentials()
    settings = load_settings()

    service_list = ["polly", "microsoft", "watson", "google", "elevenlabs"]
    voices_dict = {service: [] for service in service_list}
    tts_dict = {service: create_tts_client(service, credentials) for service in service_list}

    # Fetch voices dynamically
    for service, tts in tts_dict.items():
        voices_dict[service] = tts.get_voices()

    layout = [
        [sg.Text("Select TTS Service")],
        [sg.Combo(service_list, key="-SERVICE-", enable_events=True)],
        [sg.Listbox(values=[], size=(60, 20), select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, key="-VOICES-")],
        [sg.Button("Save Selection")]
    ]

    window = sg.Window("VoiceBroker Settings Manager", layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "-SERVICE-":
            selected_service = values["-SERVICE-"]
            voices = voices_dict[selected_service]
            voice_list = [[f"{voice['id']} - {voice['name']} ({voice.get('language_codes', ['Unknown'])[0]}, {voice['gender']})", sg.Button("Play", key=f"-PLAY-{voice['id']}-{voice.get('language_codes', ['Unknown'])[0]}-{selected_service}-")] for voice in voices]
            window["-VOICES-"].update(voice_list)
        if "Play" in event:
            _, voice_id, lang_code, selected_service = event.split("-")
            play_voice(tts_dict[selected_service], voice_id, lang_code)
        if event == "Save Selection":
            selected_service = values["-SERVICE-"]
            selected_indices = values["-VOICES-"]
            selected_voices = [voices_dict[selected_service][i] for i in range(len(voices_dict[selected_service])) if f"{voices_dict[selected_service][i]['id']} - {voices_dict[selected_service][i]['name']} ({voices_dict[selected_service][i].get('language_codes', ['Unknown'])[0]}, {voices_dict[selected_service][i]['gender']})" in selected_indices]

            voices_json = {selected_service: selected_voices}

            with open("voices.json", "w") as f:
                json.dump(voices_json, f, indent=2)

            sg.popup("Success", "Voices saved to voices.json")

    window.close()

if __name__ == "__main__":
    main()