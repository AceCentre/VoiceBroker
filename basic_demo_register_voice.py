import pythoncom
import win32com.client
import win32com.server.register
import azure.cognitiveservices.speech as speechsdk
import time
import io
import json
import logging
import winreg

def register_voice(engine, language, voiceid, name, gender, vendor):
    # Create the registry key for the voice
    key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}")

    # Set the attributes of the voice
    winreg.SetValue(key, "", winreg.REG_SZ, name)
    winreg.SetValue(key, "Language", winreg.REG_SZ, language)
    winreg.SetValue(key, "Gender", winreg.REG_SZ, gender)
    winreg.SetValue(key, "Vendor", winreg.REG_SZ, vendor)
    winreg.SetValue(key, "CLSID", winreg.REG_SZ, str(PythonTTSVoice._reg_clsid_))

    # Create the registry key for the attributes
    attr_key = winreg.CreateKey(key, "Attributes")
    winreg.SetValue(attr_key, "Language", winreg.REG_SZ, language)
    winreg.SetValue(attr_key, "Gender", winreg.REG_SZ, gender)
    winreg.SetValue(attr_key, "Vendor", winreg.REG_SZ, vendor)

    # Close the registry keys
    winreg.CloseKey(attr_key)
    winreg.CloseKey(key)

class PythonTTSVoice(win32com.client.CDispatch):
    _public_methods_ = ['Speak', 'GetOutputFormat']
    _reg_progid_ = "PythonTTSVoice.Engine"
    _reg_clsid_ = pythoncom.CreateGuid()

    def Speak(self, dwSpeakFlags, rguidFormatId, pWaveFormatEx, pTextFragList, pOutputSite):
        pass
        # Implement the Speak method here

    def GetOutputFormat(self, pTargetFmtId, pTargetWaveFormatEx, ppCoMemActualFmtId, ppCoMemActualWaveFormatEx):
        # Implement the GetOutputFormat method here
        pass 

if __name__ == '__main__':
    logging.debug("COM server registration starting...")
    win32com.server.register.UseCommandLine(PythonTTSVoice)
    register_voice("PythonTTSVoice", "409", "PythonTTSVoice1", "Python TTS Voice", "Male", "Python")

