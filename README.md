# Python SAPI Bridge for Azure Speech

This project demonstrates how to create a Python COM server that functions as a bridge between the Microsoft Azure Speech SDK and the Speech Application Programming Interface (SAPI) used by Windows. It allows applications expecting SAPI compatibility to utilize Azure's advanced neural voices for text-to-speech (TTS) functionalities.

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

# Python SAPI Bridge for Azure Speech

This project demonstrates how to create a Python COM server that functions as a bridge between the Microsoft Azure Speech SDK and the Speech Application Programming Interface (SAPI) used by Windows. It allows applications expecting SAPI compatibility to utilize Azure's advanced neural voices for text-to-speech (TTS) functionalities.

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

