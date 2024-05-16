# SAPI Bridge for Online TTS

This project demonstrates how to create a Python COM server that functions as a bridge between the Microsoft Azure Speech SDK and the Speech Application Programming Interface (SAPI) used by Windows. It allows applications expecting SAPI compatibility to utilize Azure's advanced neural voices for text-to-speech (TTS) functionalities.

Quick notes:

- voiceBrokerTTSW.py looks like it should work. Use it with demoClient for example
- But using in any SAPI system and it doesnt. We need to debug this. I think somethings missing
- Work with basic_demo_register.py and then find what is missing
- Need cred files - email me
- demoClientExtended is broken because https://github.com/willwade/py3-tts is broken on sapi in my humble opinon. But look at it as it gives interesting debug output like

```bash

2024-05-16 16:10:21,238 - DEBUG - wrap_outparam(<POINTER(ISpeechObjectTokens) ptr=0x2d193a3ea00 at 2d19302e250>)
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d194818270 at 2d19302e0d0>
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(ISpeechObjectTokens) ptr=0x2d193a3ea00 at 2d19302e250>
2024-05-16 16:10:21,238 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d194814f50 at 2d19302e1d0>)
2024-05-16 16:10:21,238 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d194814f50 at 2d19302e1d0>)
2024-05-16 16:10:21,238 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,238 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302de50>
2024-05-16 16:10:21,238 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d194814f50 at 2d19302e2d0>
2024-05-16 16:10:21,238 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,238 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,238 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d194814f50 at 2d19302de50>
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194814f50 at 2d19302e450>
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e3d0>
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e2d0>
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194814f50 at 2d19302e1d0>
2024-05-16 16:10:21,238 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948160c0 at 2d19302e1d0>)
2024-05-16 16:10:21,238 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948160c0 at 2d19302e1d0>)
2024-05-16 16:10:21,238 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,238 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,238 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e3d0>
2024-05-16 16:10:21,238 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948160c0 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948160c0 at 2d19302e3d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948160c0 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e450>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948160c0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d194814f50 at 2d19302de50>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d194816660 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d194816660 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e450>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d194816660 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d194816660 at 2d19302e450>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194816660 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194816660 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948160c0 at 2d19302e3d0>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948161e0 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948161e0 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948161e0 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948161e0 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948161e0 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948161e0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d194816660 at 2d19302e450>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948166f0 at 2d19302e5d0>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948166f0 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e5d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948161e0 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d194816780 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d194816780 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d194816780 at 2d19302e650>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d194816780 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194816780 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e650>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194816780 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948166f0 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d194814d10 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d194814d10 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d194814d10 at 2d19302e6d0>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d194814d10 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194814d10 at 2d19302e5d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e6d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d194814d10 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d194816780 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948153d0 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948153d0 at 2d19302e1d0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948153d0 at 2d19302e750>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948153d0 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948153d0 at 2d19302e650>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e5d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e750>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948153d0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d194814d10 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IEnumVARIANT) ptr=0x2d194818270 at 2d19302dbd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948153d0 at 2d19302e550>
2024-05-16 16:10:21,254 - INFO - Successfully set voice to Jessa
2024-05-16 16:10:21,254 - INFO - Text to speak: Hello, this is a test of the SAPI voice using pyttsx3.
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(ISpeechObjectTokens) ptr=0x2d193a3fa80 at 2d19302e250>)
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d194818270 at 2d19302e3d0>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302dfd0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302dfd0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948166f0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948166f0 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302dbd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948166f0 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948168a0 at 2d19302dfd0>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948168a0 at 2d19302dfd0>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948168a0 at 2d19302e750>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948168a0 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948168a0 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302dbd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e750>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948168a0 at 2d19302dfd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948166f0 at 2d19302e350>
2024-05-16 16:10:21,254 - DEBUG - wrap_outparam(<POINTER(IDispatch) ptr=0x2d1948152b0 at 2d19302e450>)
2024-05-16 16:10:21,254 - DEBUG - GetBestInterface(<POINTER(IDispatch) ptr=0x2d1948152b0 at 2d19302e450>)
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo, trying IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Does NOT implement IProvideClassInfo/IProvideClassInfo2
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d190c59100 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Default interface is {C74A3ADC-B727-4500-A84A-B526721C8B8C}
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IUnknown) ptr=0x2d1948152b0 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - GetModule(TLIBATTR(GUID={C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version=5.4, LCID=0, FLags=0x8))
2024-05-16 16:10:21,254 - DEBUG - Implements default interface from typeinfo <class 'comtypes.gen._C866CA3A_32F7_11D2_9602_00C04F8EE628_0_5_4.ISpeechObjectToken'>
2024-05-16 16:10:21,254 - DEBUG - Final result is <POINTER(ISpeechObjectToken) ptr=0x2d1948152b0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948152b0 at 2d19302e550>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeInfo) ptr=0x2d190c59100 at 2d19302dbd0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ITypeLib) ptr=0x2d190cd07b0 at 2d19302e4d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IDispatch) ptr=0x2d1948152b0 at 2d19302e450>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948168a0 at 2d19302e2d0>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(IEnumVARIANT) ptr=0x2d194818270 at 2d19302de50>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectTokens) ptr=0x2d193a3fa80 at 2d19302e250>
2024-05-16 16:10:21,254 - DEBUG - Release <POINTER(ISpeechObjectToken) ptr=0x2d1948152b0 at 2d19302e1d0>
2024-05-16 16:10:21,254 - INFO - Starting: None

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

