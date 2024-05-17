import PySimpleGUI as sg
import json
import logging
from tts_wrapper import PollyClient, PollyTTS, MicrosoftClient, MicrosoftTTS

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
    else:
        raise ValueError("Unsupported TTS service")
    return tts

def play_voice(tts, voice_id, text="Hello, world!"):
    ssml_text = tts.ssml.add(text)
    tts.speak(ssml_text, voice_id=voice_id)

def main():
    credentials = load_credentials()
    settings = load_settings()

    # Create TTS clients
    polly_tts = create_tts_client("polly", credentials)
    microsoft_tts = create_tts_client("microsoft", credentials)

    # Fetch voices (simulated with settings)
    polly_voices = settings.get("polly", [])
    microsoft_voices = settings.get("microsoft", [])

    service_list = ["polly", "microsoft"]
    voices_dict = {
        "polly": polly_voices,
        "microsoft": microsoft_voices
    }
    tts_dict = {
        "polly": polly_tts,
        "microsoft": microsoft_tts
    }

    layout = [
        [sg.Text("Select TTS Service")],
        [sg.Combo(service_list, key="-SERVICE-", enable_events=True)],
        [sg.Listbox(values=[], size=(60, 20), select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, key="-VOICES-")],
        [sg.Button("Save Selection")]
    ]

    window = sg.Window("TTS Voice Selector", layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "-SERVICE-":
            selected_service = values["-SERVICE-"]
            voices = voices_dict[selected_service]
            voice_list = [[f"{voice['voiceid']} - {voice['name']} ({voice['language']}, {voice['gender']})", sg.Button("Play", key=f"-PLAY-{voice['voiceid']}-")] for voice in voices]
            window.extend_layout(window, [[sg.Column(voice_list)]])
        if "Play" in event:
            selected_service = values["-SERVICE-"]
            voice_id = event.split("-")[2]
            play_voice(tts_dict[selected_service], voice_id)
        if event == "Save Selection":
            selected_service = values["-SERVICE-"]
            selected_indices = values["-VOICES-"]
            selected_voices = [voices_dict[selected_service][i] for i in range(len(voices_dict[selected_service])) if f"{voices_dict[selected_service][i]['voiceid']} - {voices_dict[selected_service][i]['name']} ({voices_dict[selected_service][i]['language']}, {voices_dict[selected_service][i]['gender']})" in selected_indices]

            voices_json = {selected_service: selected_voices}

            with open("voices.json", "w") as f:
                json.dump(voices_json, f, indent=2)

            sg.popup("Success", "Voices saved to voices.json")

    window.close()

if __name__ == "__main__":
    main()