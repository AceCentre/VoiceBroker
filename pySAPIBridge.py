import pythoncom
import win32com.client
import azure.cognitiveservices.speech as speechsdk
import time
import json
import pygame
import json
import win32com.server.register

def load_credentials(file_path):
    try:
        with open(file_path, 'r') as file:
            return json.load(file)
    except FileNotFoundError:
        print("The specified credentials file was not found.")
        return None
    except json.JSONDecodeError:
        print("Error decoding JSON from the credentials file.")
        return None


class PythonTTSVoice():
    _public_methods_ = ['Speak', 'Pause', 'Resume', 'GetVoices', 'SetVoice', 'SetInterest', 'WaitForNotifyEvent']
    _reg_progid_ = "PythonTTSVoice"
    _reg_clsid_ = pythoncom.CreateGuid()

    def __init__(self, credentials):
        # Initialize the COM library within the class
        pythoncom.CoInitialize()
        # Create an instance of the SAPI SpVoice COM object
        self.sp_voice = win32com.client.Dispatch("SAPI.SpVoice")
        microsoft_creds = credentials['Microsoft']
        self.speech_config = speechsdk.SpeechConfig(subscription=microsoft_creds['TOKEN'], region=microsoft_creds['region'])
        self.stream = speechsdk.audio.PullAudioOutputStream()
        self.audio_config = speechsdk.audio.AudioConfig(stream=self.stream)
        self.speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=self.speech_config, audio_config=self.audio_config)
        self.event_interests = {}
        self.current_voice = 'en-US-JessaNeural'
        pygame.mixer.init()

    def Speak(self, text):
        self.speech_config.request_word_level_timestamps()
        stream = speechsdk.audio.AudioOutputStream.create_pull_audio_output_stream()
        audio_config = speechsdk.audio.AudioConfig(stream=stream)
        synthesizer = speechsdk.SpeechSynthesizer(speech_config=self.speech_config, audio_config=audio_config)
        synthesis_future = synthesizer.speak_text_async(text)
        result = synthesis_future.get()

        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            json_result = json.loads(result.json)
            words_info = json_result.get('NBest')[0].get('Words')
            self.words_info = words_info  # Store word timing data for event synchronization
            self.audio_stream = stream
            self.play_audio(stream, words_info)
        elif result.reason == speechsdk.ResultReason.Canceled:
            print("Speech synthesis canceled: {}".format(result.cancellation_details.reason))

    def play_audio(self, stream, words_info):
        # Detach the stream and read all data into a byte buffer
        stream.detach()
        data = self.stream.read_all()
        sound_file = io.BytesIO(data)
        sound = pygame.mixer.Sound(file=sound_file)
        
        # Start playback and handle timing events
        self.playback = sound.play()
        start_time = pygame.time.get_ticks()  # Get the current tick count

        # Monitor playback and trigger events based on word timings
        while pygame.mixer.get_busy():
            current_time = pygame.time.get_ticks() - start_time
            for word in words_info:
                if 'start_time' not in word:
                    if word['Offset'] <= current_time <= word['Offset'] + word['Duration']:
                        print(f"Word spoken: {word['Word']}")
                        word['start_time'] = current_time  # Mark as handled

            pygame.time.delay(100)  # Check every 100 milliseconds
    
    def Pause(self):
        pygame.mixer.pause()  # This pauses all sounds in the mixer

    def Resume(self):
        pygame.mixer.unpause()  # This resumes all paused sounds

    def GetVoices(self):
        voices = self.speech_synthesizer.get_voices_list()
        return [{'Name': voice.short_name, 'Locale': voice.locale, 'Gender': voice.gender} for voice in voices]
    
    def SetVoice(self, voice_name):
        # Setting a voice directly if it exists in the fetched voices
        for voice in self.speech_synthesizer.get_voices_list():
            if voice.short_name == voice_name:
                self.speech_synthesizer.properties[speechsdk.PropertyId.SpeechServiceConnection_SynthVoice] = voice_name
                return f"Voice set to {voice_name}"
        return "Voice not found"

    def SetInterest(self, event_id, enabled):
        self.event_interests[event_id] = enabled
    
    def WaitForNotifyEvent(self, timeout):
        # This would be an implementation where you monitor the audio playback and check for the next event
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.event_triggered:
                return self.event_data
            time.sleep(0.1)  # Sleep briefly to avoid high CPU usage
        return None

        
if __name__ == '__main__':
    credentials_path = 'credentials.json'
    credentials = load_credentials(credentials_path)
    if credentials:
        tts_voice = PythonTTSVoice(credentials)
        print("COM server registration starting...")
        win32com.server.register.UseCommandLine(PythonTTSVoice, "--register")
    else:
        print("Failed to load credentials. Check your credentials file and try again.")