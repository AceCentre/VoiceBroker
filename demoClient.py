import win32com.client

def test_com_server():
    # Create a COM object instance from the PythonTTSVoice server
    server = win32com.client.Dispatch("VoiceBroker.Application")

    # Retrieve and print the list of available voices
    voices = server.GetVoices()
    print(f"Available Voices: {voices}")

    # Speak a sample text
    server.Speak("Hello world")

if __name__ == "__main__":
    test_com_server()
