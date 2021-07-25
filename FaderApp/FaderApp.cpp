// FaderApp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

//for audio
#include <iostream>
#include <mmdeviceapi.h>
#include <combaseapi.h>
#include <audiopolicy.h>
#include <psapi.h>
#include <Audioclient.h>

#include "SerialPort.hpp"
#include <string> 


int selected = 0;
int count = 0;

ISimpleAudioVolume* activeAudioVolume;
IAudioSessionEnumerator* sessionEnumerator;
char currentAppName[MAX_PATH];
int lastVolume = 0;


void refreshtSessionEnumerator() {
    HRESULT hr;
    IMMDevice* speaker = NULL;
    IMMDeviceEnumerator* deviceEnumerator = NULL;

    hr = CoInitializeEx(NULL, COINITBASE_MULTITHREADED); //init external communication
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, 1, IID_PPV_ARGS(&deviceEnumerator));
    if (FAILED(hr)) {
        printf("FAIL");
    }

    hr = deviceEnumerator->GetDefaultAudioEndpoint(eRender, eMultimedia, &speaker); //get the default device
    if (FAILED(hr)) { //check if it worked
        printf("Fail");
    }

    //get the sessions from this device
    //we want an AudioSession Manager2
    IAudioSessionManager2* pSessionManager;
    hr = speaker->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, NULL, (void**)&pSessionManager);
    if (FAILED(hr)) {
        printf("FAILED");
    }

    //get an enumerator for all active sessions
    IAudioSessionEnumerator* sessionEnum = NULL;
    hr = pSessionManager->GetSessionEnumerator((IAudioSessionEnumerator**)&sessionEnum);
    if (FAILED(hr)) { //check if it worked
        printf("Fail");
    }
    sessionEnumerator = sessionEnum;
}

int getSessionCount() {
    refreshtSessionEnumerator();
    HRESULT hr;
    int count = 0;
    hr = sessionEnumerator->GetCount((int*)&count);
    if (FAILED(hr)) {
        printf("Fail");
    }
    return count;
}

void getAppName(int appID, char* frName) {
    HRESULT hr;
    IAudioSessionControl2* controlSession = NULL;
    hr = sessionEnumerator->GetSession(appID, (IAudioSessionControl**)&controlSession);
    if (FAILED(hr)) {
        return;
    }

    IAudioSessionControl* ctrl = NULL;
    IAudioSessionControl2* ctrl2 = NULL;
    DWORD processId = 0;

    hr = sessionEnumerator->GetSession(appID, &ctrl);
    if (FAILED(hr)) {
        printf("FAIL");
        return;
    }
    hr = ctrl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&ctrl2);
    if (FAILED(hr))
    {
        printf("FAIL");
        return;
    }

    hr = ctrl2->GetProcessId(&processId);
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, false, processId);
    DWORD value = MAX_PATH;
    char buffer[MAX_PATH];
    //int ret = QueryFullProcessImageName(hProcess, 0, (LPWSTR)buffer, &value);
    DWORD ret = GetProcessImageFileNameA(hProcess, (LPSTR)buffer, MAX_PATH);
    if (ret == 0) {
        return;
    }
    CloseHandle(hProcess);


    char* cp = reinterpret_cast<char*>(buffer); //get a pointer to the buffer
    cp += ret; //shift it on the last char now point on the NULL
    int loopCount = 0;
    while (*cp != 0x5c) {
        cp--;
        if (*cp == '.') {
            *cp = '\0';
        }
        loopCount++;
    }
    cp++;
    for (int j = 0; j < loopCount - 1; j++) {
        frName[j] = cp[j];
    }
}

float getVolume() {
    HRESULT hr;
    float val;
    hr = activeAudioVolume->GetMasterVolume(&val);
    if (FAILED(hr)) { //check if it worked
        printf("ERROR");
        return 0;
    }
    return val;    
}

void refreshSimpleAudioVolume(int id) {
    HRESULT hr;
    IAudioSessionControl2* controlSession = NULL;
    hr = sessionEnumerator->GetSession(id, (IAudioSessionControl**)&controlSession);
    if (FAILED(hr)) {
        printf("ERROR");
        return;
    }

    //get the VolumeController
    ISimpleAudioVolume* audioVolume;
    GUID guid;
    controlSession->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&audioVolume); //get the ISimpleAudioVolume of the current session
    activeAudioVolume = audioVolume;
}

void setVolume(float val) {
    HRESULT hr;
    GUID guid;
    hr = activeAudioVolume->SetMasterVolume(val, (GUID*)&guid);
    if (FAILED(hr)) {
        printf("ERROR");
        return;
    }
}

BOOL isMute() {
    HRESULT hr;
    BOOL isMute;
    hr = activeAudioVolume->GetMute((BOOL*)&isMute);
    if (FAILED(hr)) {
        std::cerr << "ERR";
        return false;
    }
    return isMute;
}

void toggleMute() {
    HRESULT hr;
    GUID guid;

    hr = activeAudioVolume->SetMute((BOOL)!isMute(), (GUID*)&guid);
    if (FAILED(hr)) {
        std::cerr << "ERROR TOGGLE MUTE";
    }
}

void printNames() {
    int count = getSessionCount();
    for (int i = 0; i < count; i++) {
        char frName[MAX_PATH];
        getAppName(i, frName);
        std::cout << i;
        printf("%s", frName);
        std::cout << "\n";
    }
}

void setNewApp() {
    refreshtSessionEnumerator();
    refreshSimpleAudioVolume(selected);
    lastVolume = getVolume();
    getAppName(selected, currentAppName);
}

void sendSerial(float vol, int muted, char* name, SerialPort* arduino) {
    std::string send;
    
    int sendVal = (int)(vol * 100);
    std::string volumeString = std::to_string(sendVal);
    std::string muteString = std::to_string(muted);
    send.append("v");
    send.append(volumeString);
    send.append(";");
    send.append("m");
    send.append(muteString);
    send.append(";");
    send.append("a");
    send.append(std::to_string(selected));
    send.append(":");
    send.append(name);
    send.append(";");
    const char* sendString = send.c_str();
    bool hasWritten = arduino->writeSerialPort(sendString, send.length());
    if (!hasWritten) std::cerr << "Data was not written" << std::endl;
}

void send(SerialPort* arduino) {
    int volume = (int)(getVolume() * 100);
    std::string send;
    send.append("v;");
    send.append(std::to_string(volume));
    send.append(";a;");
    send.append(currentAppName);
    send.append("\n");
    std::cout << send;
    const char* sendString = send.c_str();
    arduino->writeSerialPort(sendString, send.length());
}

int main()
{
    setNewApp();
    printNames();
    const char* portName = "\\\\.\\COM6";
    SerialPort* arduino;
    arduino = new SerialPort(portName);
    std::cout << "Is connected: " << arduino->isConnected() << std::endl;
    
    while (true) {
        char rec[1];
        arduino->readSerialPort(rec, 1);
        while (rec[0] != 'b') {
            arduino->readSerialPort(rec, 1);
            Sleep(50);
        }
        arduino->readSerialPort(rec, 1);
        if (rec[0] == '1') {
            selected++;
            if (selected >= getSessionCount()) {
                selected = 0;
            }
            setNewApp();
            std::cout << "\n";
            std::cout << selected;
            printf("%s", currentAppName);
            std::cout << "\n\n";
            printNames();
            send(arduino);  
        }
        else if (rec[0] == '2') {
            selected--;
            if (selected <= -1) {
                selected = getSessionCount() - 1;
            }
            setNewApp();
            std::cout << "\n";
            std::cout << selected;
            printf("%s", currentAppName);
            std::cout << "\n\n";
            printNames();
            send(arduino);
        }
        else if (rec[0] == '0') {
            arduino->readSerialPort(rec, 1);
            float vol = ((float)(int)(rec[0]) / 100);
            setVolume(vol);
            std::cout << (int)rec[0] << '\n';
        }
    }   
}

