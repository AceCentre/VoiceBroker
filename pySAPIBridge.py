import pythoncom
import win32com.client
import azure.cognitiveservices.speech as speechsdk
import time
import json
import simpleaudio

class PythonTTSVoice(win32com.client.Dispatch):
    _public_methods_ = ['Speak', 'Pause', 'Resume', 'GetVoices', 'SetVoice', 'SetInterest', 'WaitForNotifyEvent']
    _reg_progid_ = "Python.TTSVoice"
    _reg_clsid_ = pythoncom.CreateGuid()

    def __init__(self, api_key, region):
        self.speech_config = speechsdk.SpeechConfig(subscription=api_key, region=region)
        self.audio_config = speechsdk.audio.AudioConfig(use_default_microphone=False)  # Corrected as no microphone needed
        self.speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=self.speech_config, audio_config=None)
        self.event_interests = {}
        self.current_voice = 'en-US-JessaNeural'

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
            self.words_info = words_info
            self.audio_stream = stream
            self.play_audio(stream)
        elif result.reason == speechsdk.ResultReason.Canceled:
            print("Speech synthesis canceled: {}".format(result.cancellation_details.reason))

    def play_audio(self, stream):
        stream.detach()
        data = stream.read_all()
        self.playback = simpleaudio.WaveObject(data, 1, 2, 16000).play()
        self.playback.wait_done()

    
    def Pause(self):
        if self.playback:
            self.playback.stop()
    
    def Resume(self):
        if self.audio_stream:
            self.play_audio(self.audio_stream)

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
    print("Registering COM server...")
    win32com.server.register.UseCommandLine(PythonTTSVoice, "--register")