import pythoncom
import win32com.client

def main():
    pythoncom.CoInitialize()
    try:
        sapi_voice = win32com.client.Dispatch("SAPI.SpVoice")
        for voice in sapi_voice.GetVoices():
            if "VoiceBroker" in voice.GetDescription():
                sapi_voice.Voice = voice
                break
        sapi_voice.Speak("Hello, this is a test of the Python TTS voice.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    main()
