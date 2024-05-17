import PySimpleGUI as sg
import json
import logging
from tts_wrapper import PollyClient, PollyTTS, MicrosoftClient, MicrosoftTTS

def create_tts_client(service, settings):
    if service == "polly":
        creds = settings.get('Polly', {})
        client = PollyClient(credentials=(creds.get('aws_key_id'), creds.get('aws_secret_access_key')))
        tts = PollyTTS(client=client)
    elif service == "microsoft":
        creds = settings.get('Microsoft', {})
        client = MicrosoftClient(credentials=creds.get('subscription_key'), region=creds.get('region'))
        tts = MicrosoftTTS(client=client)
    else:
        raise ValueError("Unsupported TTS service")
    return tts

def load_settings():
    try:
        with open('settings.json', 'r') as f:
            settings = json.load(f)
        logging.debug("Settings loaded from file.")
        return settings
    except FileNotFoundError:
        logging.error("Settings file not found. Ensure 'settings.json' is in the correct location.")
        return {}

def main():
    settings = load_settings()
    polly_tts = create_tts_client("polly", settings)
    microsoft_tts = create_tts_client("microsoft", settings)

    polly_voices = polly_tts.get_voices()
    microsoft_voices = microsoft_tts.get_voices()

    service_list = ["polly", "microsoft"]
    voices_dict = {
        "polly": polly_voices,
        "microsoft": microsoft_voices
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
            voice_list = [f"{voice['Id']} - {voice['Name']} ({voice['LanguageCode']}, {voice['Gender']})" for voice in voices]
            window["-VOICES-"].update(voice_list)
        if event == "Save Selection":
            selected_service = values["-SERVICE-"]
            selected_indices = values["-VOICES-"]
            selected_voices = [voices_dict[selected_service][i] for i in range(len(voices_dict[selected_service])) if voices_dict[selected_service][i]['Id'] in [item.split(' - ')[0] for item in selected_indices]]

            voices_json = {selected_service: selected_voices}

            with open("voices.json", "w") as f:
                json.dump(voices_json, f, indent=2)

            sg.popup("Success", "Voices saved to voices.json")

    window.close()

if __name__ == "__main__":
    main()