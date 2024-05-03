import win32com.client

def test_com_server():
    server = win32com.client.Dispatch("PythonTTSVoice.Application")
    server.Speak("Hello world")

if __name__ == "__main__":
    test_com_server()
