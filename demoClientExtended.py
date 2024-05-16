import pyttsx3
import logging

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
    return voices

def set_voice(engine, voice_name):
    voices = engine.getProperty('voices')
    for voice in voices:
        if voice.name == voice_name:
            engine.setProperty('voice', voice.id)
            return True
    return False

def on_start(name):
    logging.info(f"Starting: {name}")

def on_word(name, location, length):
    logging.info(f"Word: {name}, {location}, {length}")

def on_end(name, completed):
    logging.info(f"Finished: {name}, {completed}")

def main():
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

    try:
        # Initialize the TTS engine
        engine = pyttsx3.init()
        logger.info("TTS engine initialized.")

        # List available voices
        logger.info("Available voices:")
        voices = list_voices(engine)

        # Set the desired voice
        voice_name = "Jessa"  # Example voice name, change as needed
        if set_voice(engine, voice_name):
            logger.info(f"Successfully set voice to {voice_name}")
        else:
            logger.error(f"Could not find voice named {voice_name}")
            return

        # Connect event callbacks
        engine.connect('started-utterance', on_start)
        engine.connect('started-word', on_word)
        engine.connect('finished-utterance', on_end)

        # Speak some text using the chosen voice
        text_to_speak = "Hello, this is a test of the SAPI voice using pyttsx3."
        logger.info(f"Text to speak: {text_to_speak}")
        engine.say(text_to_speak)

        # Run and wait for the engine to complete processing all commands
        engine.runAndWait()
        logger.info("Finished speaking.")
    except Exception as e:
        logger.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
