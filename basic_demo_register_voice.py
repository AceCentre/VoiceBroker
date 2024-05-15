import pythoncom
import win32com.client
import win32com.server.register
import azure.cognitiveservices.speech as speechsdk
import time
import io
import json
import logging
import winreg


def register_voice(voiceid, name, language, gender, vendor):
    # Paths for 64-bit and 32-bit applications
    paths = [
        f"SOFTWARE\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}",
        f"SOFTWARE\\WOW6432Node\\Microsoft\\SPEECH\\Voices\\Tokens\\{voiceid}"
    ]

    for path in paths:
        # Create the registry key for the voice
        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path)

        # Set the attributes of the voice
        winreg.SetValue(key, "", winreg.REG_SZ, name)
        winreg.SetValue(key, "409", winreg.REG_SZ, language)
        winreg.SetValue(key, "CLSID", winreg.REG_SZ, "{4DFFD59B-4DF3-4366-B053-DFF9BE002EFB}")

        # Create the registry key for the attributes
        attr_key = winreg.CreateKey(key, "Attributes")
        winreg.SetValue(attr_key, "Language", winreg.REG_SZ, language)
        winreg.SetValue(attr_key, "Gender", winreg.REG_SZ, gender)
        winreg.SetValue(attr_key, "Vendor", winreg.REG_SZ, vendor)
        winreg.SetValue(attr_key, "Name", winreg.REG_SZ, name)

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

