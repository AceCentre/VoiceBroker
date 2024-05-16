import pyttsx3

def list_voices(engine):
    voices = engine.getProperty('voices')
    for i, voice in enumerate(voices):
        print(f"Voice {i+1}:")
        print(f"  ID: {voice.id}")
        print(f"  Name: {voice.name}")
        print(f"  Languages: {voice.languages}")
        print(f"  Gender: {voice.gender}")
        print(f"  Age: {voice.age}")
        print()

def set_voice(engine, voice_name):
    voices = engine.getProperty('voices')
    for voice in voices:
        if voice.name == voice_name:
            engine.setProperty('voice', voice.id)
            return True
    return False

def main():
    # Initialize the TTS engine
    engine = pyttsx3.init()

    # List available voices
    print("Available voices:")
    list_voices(engine)

    # Set the desired voice
    voice_name = "Jessa"
    if set_voice(engine, voice_name):
        print(f"Successfully set voice to {voice_name}")
    else:
        print(f"Could not find voice named {voice_name}")

    # Speak some text using the chosen voice
    text_to_speak = "Hello, this is a test of the SAPI voice using pyttsx3."
    engine.say(text_to_speak)
    engine.runAndWait()

if __name__ == "__main__":
    main()
