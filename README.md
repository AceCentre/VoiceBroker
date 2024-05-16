# SAPI Bridge for Online TTS

This project demonstrates how to create a Python COM server that functions as a bridge between the Microsoft Azure Speech SDK and the Speech Application Programming Interface (SAPI) used by Windows. It allows applications expecting SAPI compatibility to utilize Azure's advanced neural voices for text-to-speech (TTS) functionalities.

## Quick notes:

- voiceBroker.py looks like it should work. Use it with demoClient for example. This is the code we want to work from
- voiceBroker-MSVanilla.py is not using tts-wrapper. Just for Azure, The code is a little older - i dare the creds might not work.
- But using in any SAPI system and it doesnt. We need to debug this. I think somethings missing (e.g try it in https://www.cross-plus-a.com/balabolka.htm)
- Work with basic_demo_register.py and then find what is missing
- You need the correct credentials.json
- demoClientExtended - try it by all means but I have little faith in the onWord event handling being passed through correctly. 


Here's what I think is going on. I think we have largely followed the details here. https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms720163(v=vs.85)
We register the Engine and com service. That definitely works and we can get speak working with demoClient.py. But trying it any proper SAPI application doesnt. We are finding it hard to debug this.
I have a feeling we have methods that arent implemented or speech isnt being called correctly. Im struggling to figure out from the MS docs whats going on. 

**Note:** look at requirements. we are using py3.11.4 and using our own forks of py3-tts - you'll see why. 

```

## Aims

- [x] Basic setting as a COM Service wih some functions
- [] Correctly register in the registry
- [] First work on Azure (why? Because I know we can get word level timing - it should help for callbacks)
- [] Extend this whole code for other libraries (Abstract the class out)
- [] Allow for configuration by a seperate app to allow for choice of voices etc


## Features

- Utilizes Azure Cognitive Services Speech SDK for high-quality neural voices.
- Implements SAPI-like methods such as `Speak`, `Pause`, `Resume`, and voice selection.
- Supports word-level timestamp events for synchronized operations.
- Designed to work with any SAPI-compatible software on Windows.

## Prerequisites

Before you start using this project, ensure you have the following:
- Python 3.7 or higher
- Azure subscription and Speech service created on Azure portal
- Required Python packages: `azure-cognitiveservices-speech`, `pythoncom`, `win32com.client`, `simpleaudio`

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/yourgithubusername/python-sapi-azure.git
   cd python-sapi-azure
   ```
2, **Install Dependencies**

    ```bash
    pip install azure-cognitiveservices-speech pythoncom pywin32 simpleaudio
    ```

3. **Configure Azure API Key and Region**

Set up your Azure API key and region in the configuration section of the script or as environment variables.

## Usage

To run the server and register it as a COM object on your Windows machine, execute:

```bash
python PythonTTSVoice.py --register
```

After registration, your Python COM server will be available for use by any SAPI-compatible software.

## Methods Implemented

* Speak(text): Synthesizes speech from text using Azure's TTS.
* Pause(): Pauses the ongoing speech synthesis.
* Resume(): Resumes the paused speech synthesis.
* GetVoices(): Lists available voices from Azure.
* SetVoice(voice_name): Sets the current voice to the specified one.
* SetInterest(event_id, enabled): Subscribes or unsubscribes from specific speech events.
* WaitForNotifyEvent(timeout): Waits for a subscribed event to occur within a specified timeout period.

## Contributing

Contributions are MORE than welcome. This code is not tested whatsoever. Please improve

