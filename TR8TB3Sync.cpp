#include <Windows.h>
#include <initguid.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <functiondiscoverykeys_devpkey.h>
#include <avrt.h>
#include <strsafe.h>
#include <comdef.h>
#define _USE_MATH_DEFINES
#include <math.h>


#include <string>
#include <vector>
#include <stdexcept>
using namespace std;

#include "resource.h"

#include "debug.h"

const DWORD SAMPLERATE = 96000;

UINT tr8_index_in = UINT_MAX;
UINT tb3_index_in = UINT_MAX;
UINT tr8_index_out = UINT_MAX;
UINT tb3_index_out = UINT_MAX;
UINT tr8_midi_id_in = UINT_MAX;
UINT tb3_midi_id_in = UINT_MAX;
UINT tr8_midi_id_out = UINT_MAX;
UINT tb3_midi_id_out = UINT_MAX;

vector<wstring> devices;

_COM_SMARTPTR_TYPEDEF(IMMDeviceCollection, __uuidof(IMMDeviceCollection));
_COM_SMARTPTR_TYPEDEF(IMMDeviceEnumerator, __uuidof(IMMDeviceEnumerator));
_COM_SMARTPTR_TYPEDEF(IMMDevice, __uuidof(IMMDevice));
_COM_SMARTPTR_TYPEDEF(IAudioClient, __uuidof(IAudioClient));
_COM_SMARTPTR_TYPEDEF(IAudioRenderClient, __uuidof(IAudioRenderClient));
_COM_SMARTPTR_TYPEDEF(IAudioCaptureClient, __uuidof(IAudioCaptureClient));

IMMDeviceCollectionPtr deviceCollectionOut;
IMMDeviceCollectionPtr deviceCollectionIn;
IAudioClientPtr RenderAudioClient;
IAudioRenderClientPtr RenderClient;

IAudioClientPtr CaptureAudioClient[2];
IAudioCaptureClientPtr CaptureClient[2];

HANDLE ShutdownEvent;
HANDLE AudioSamplesReadyEvent;

HANDLE RenderThread;
UINT32 RenderBufferSize;

// リングバッファ
const int RING_BUF_SIZE = 9600;
int ring_buf[2][RING_BUF_SIZE * 2];
int rec_pos[2];
int play_pos;
int rec_play_diff[2];

// MIDI
HMIDIIN hmi;
HMIDIOUT hmo;

WORD wBitsPerSample = 32;

INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
bool GetDeviceName(IMMDeviceCollection *DeviceCollection, UINT DeviceIndex, LPWSTR deviceName, size_t size);
void create_event();
DWORD WINAPI WASAPIRenderThread(LPVOID Context);
void stop();
void rec_data();
void CALLBACK MidiInProc(HMIDIIN hMidiIn, UINT wMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2);


template <class T> inline void SafeRelease(T *ppT)
{
	if (*ppT)
	{
		(*ppT)->Release();
		*ppT = NULL;
	}
}

class win32_error : runtime_error
{
private:
	HRESULT _hr;

public:
	win32_error(const string& message, HRESULT hr) : runtime_error(message), _hr(hr) {}
};

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	// 出力デバイス列挙
	// デバイス列挙
	IMMDeviceEnumeratorPtr deviceEnumerator;
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&deviceEnumerator));
	if (FAILED(hr)) {
		TRACE(L"Unable to instantiate device enumerator %x", hr);
		return 0;
	}

	hr = deviceEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &deviceCollectionOut);
	if (FAILED(hr))	{
		TRACE(L"Unable to retrieve device collection %x", hr);
		return 0;
	}

	UINT deviceCount;
	hr = deviceCollectionOut->GetCount(&deviceCount);
	if (FAILED(hr))	{
		TRACE(L"Unable to get device collection length %x", hr);
		return 0;
	}

	for (UINT i = 0; i < deviceCount; i += 1)
	{
		wchar_t deviceName[128];

		if (!GetDeviceName(deviceCollectionOut, i, deviceName, sizeof(deviceName))) {
			deviceName[0] = '\0';
		}

		devices.push_back(deviceName);

		if (wcscmp(deviceName, L"OUT (TR-8)") == 0) {
			tr8_index_out = i;

		}
		else if (wcscmp(deviceName, L"OUT (TB-3)") == 0) {
			tb3_index_out = i;
		}
	}

	// TR-8とTB-3の入力デバイス確認
	hr = deviceEnumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &deviceCollectionIn);
	if (FAILED(hr))	{
		TRACE(L"Unable to retrieve device collection %x", hr);
		return 0;
	}

	hr = deviceCollectionIn->GetCount(&deviceCount);
	if (FAILED(hr))	{
		TRACE(L"Unable to get device collection length %x", hr);
		return 0;
	}

	for (UINT i = 0; i < deviceCount; i += 1)
	{
		wchar_t deviceName[128];

		if (!GetDeviceName(deviceCollectionIn, i, deviceName, sizeof(deviceName))) {
			deviceName[0] = '\0';
		}
		TRACE(L"%d:\t%s\n", i, deviceName);

		if (wcscmp(deviceName, L"IN MIX (TR-8)") == 0) {
			tr8_index_in = i;

		}
		else if (wcscmp(deviceName, L"IN (TB-3)") == 0) {
			tb3_index_in = i;
		}
	}

	// TR-8とTB-3のMIDI INデバイス確認
	UINT devMidiInNum = midiInGetNumDevs();

	MIDIINCAPS midiincaps;
	for (UINT i = 0; i < devMidiInNum; i++) {
		midiInGetDevCaps(i, &midiincaps, sizeof(MIDIINCAPS));
		if (wcscmp(midiincaps.szPname, L"TR-8") == 0) {
			tr8_midi_id_in = i;
		}
		else if (wcscmp(midiincaps.szPname, L"TB-3") == 0) {
			tb3_midi_id_in = i;
		}
	}

	// TR-8とTB-3のMIDI OUTデバイス確認
	UINT devMidiOutNum = midiOutGetNumDevs();

	MIDIOUTCAPS midioutcaps;
	for (UINT i = 0; i < devMidiOutNum; i++) {
		midiOutGetDevCaps(i, &midioutcaps, sizeof(MIDIINCAPS));
		if (wcscmp(midioutcaps.szPname, L"TR-8") == 0) {
			tr8_midi_id_out = i;
		}
		else if (wcscmp(midioutcaps.szPname, L"TB-3") == 0) {
			tb3_midi_id_out = i;
		}
	}

	// イベント作成
	create_event();

	// ダイアログ表示
	HWND hDlg = CreateDialog(hInstance, MAKEINTRESOURCE(IDD_DLG_MAIN), NULL, DialogProc);

	//メッセージループ
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	stop();

	return (int)(msg.wParam);
}

void create_event()
{
	// イベント作成
	//
	//  Create our shutdown and samples ready events- we want auto reset events that start in the not-signaled state.
	//
	ShutdownEvent = CreateEventEx(NULL, NULL, 0, EVENT_MODIFY_STATE | SYNCHRONIZE);
	if (ShutdownEvent == NULL) {
		throw win32_error("Unable to create shutdown event", GetLastError());
	}

	AudioSamplesReadyEvent = CreateEventEx(NULL, NULL, 0, EVENT_MODIFY_STATE | SYNCHRONIZE);
	if (AudioSamplesReadyEvent == NULL)	{
		throw win32_error("Unable to create samples ready event", GetLastError());
	}
}

void init_capture(int i, int index)
{
	HRESULT hr;

	// 出力デバイス準備
	IMMDevicePtr device;
	hr = deviceCollectionIn->Item(index, &device);
	if (FAILED(hr)) {
		throw win32_error("Unable to retrieve device", hr);
	}

	// AudioClient準備

	//
	//  Now activate an IAudioClient object on our preferred endpoint and retrieve the mix format for that endpoint.
	//
	hr = device->Activate(__uuidof(IAudioClient), CLSCTX_INPROC_SERVER, NULL, reinterpret_cast<void **>(&CaptureAudioClient[i]));
	if (FAILED(hr))	{
		throw win32_error("Unable to activate audio client", hr);
	}

	// デバイスのレイテンシ取得
	REFERENCE_TIME DefaultDevicePeriod;
	REFERENCE_TIME MinimumDevicePeriod;
	hr = CaptureAudioClient[i]->GetDevicePeriod(&DefaultDevicePeriod, &MinimumDevicePeriod);
	if (FAILED(hr))	{
		throw win32_error("Unable to get device period", hr);
	}

	WAVEFORMATEXTENSIBLE WaveFormat;
	WaveFormat.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
	WaveFormat.Format.nChannels = 2;
	WaveFormat.Format.nSamplesPerSec = SAMPLERATE;
	WaveFormat.Format.wBitsPerSample = 32;
	WaveFormat.Format.nBlockAlign = WaveFormat.Format.wBitsPerSample / 8 * WaveFormat.Format.nChannels;
	WaveFormat.Format.nAvgBytesPerSec = WaveFormat.Format.nSamplesPerSec * WaveFormat.Format.nBlockAlign;
	WaveFormat.Format.cbSize = 22;
	WaveFormat.dwChannelMask = SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT;
	WaveFormat.Samples.wValidBitsPerSample = 24;
	WaveFormat.SubFormat = KSDATAFORMAT_SUBTYPE_PCM;

	//  Initialize WASAPI in event driven mode, associate the audio client with our samples ready event handle, retrieve 
	//  a capture client for the transport, create the capture thread and start the audio engine.
	//
	hr = CaptureAudioClient[i]->Initialize(AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_STREAMFLAGS_NOPERSIST, DefaultDevicePeriod, DefaultDevicePeriod, (WAVEFORMATEX*)&WaveFormat, NULL);
	if (FAILED(hr))	{
		throw win32_error("Unable to initialize audio client", hr);
	}

	hr = CaptureAudioClient[i]->GetService(IID_PPV_ARGS(&CaptureClient[i]));
	if (FAILED(hr)) {
		throw win32_error("Unable to get new capture client", hr);
	}
}

void init_midi()
{
	if (tr8_midi_id_in != UINT_MAX && tb3_midi_id_out != UINT_MAX) {
		// MIDI OUTオープン
		MMRESULT mmret = midiOutOpen(&hmo, tb3_midi_id_out, NULL, NULL, CALLBACK_NULL);
		if (mmret != MMSYSERR_NOERROR) {
			return;
		}

		// MIDI INオープン
		mmret = midiInOpen(&hmi, tr8_midi_id_in, (DWORD_PTR)MidiInProc, NULL, CALLBACK_FUNCTION);
		if (mmret != MMSYSERR_NOERROR) {
			return;
		}

		midiInStart(hmi);
	}
}

void ready(UINT index)
{
	// MIDI準備
	init_midi();

	HRESULT hr;

	// 出力デバイス準備
	IMMDevicePtr device;
	hr = deviceCollectionOut->Item(index, &device);
	if (FAILED(hr)) {
		throw win32_error("Unable to retrieve device", hr);
	}

	// AudioClient準備

	//
	//  Now activate an IAudioClient object on our preferred endpoint and retrieve the mix format for that endpoint.
	//
	hr = device->Activate(__uuidof(IAudioClient), CLSCTX_INPROC_SERVER, NULL, reinterpret_cast<void **>(&RenderAudioClient));
	if (FAILED(hr))	{
		throw win32_error("Unable to activate audio client", hr);
	}

	// デバイスのレイテンシ取得
	REFERENCE_TIME DefaultDevicePeriod;
	REFERENCE_TIME MinimumDevicePeriod;
	hr = RenderAudioClient->GetDevicePeriod(&DefaultDevicePeriod, &MinimumDevicePeriod);
	if (FAILED(hr))	{
		throw win32_error("Unable to get device period", hr);
	}

	WAVEFORMATEXTENSIBLE WaveFormat;
	WaveFormat.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
	WaveFormat.Format.nChannels = 2;
	WaveFormat.Format.nSamplesPerSec = SAMPLERATE;
	WaveFormat.Format.wBitsPerSample = wBitsPerSample;
	WaveFormat.Format.nBlockAlign = WaveFormat.Format.wBitsPerSample / 8 * WaveFormat.Format.nChannels;
	WaveFormat.Format.nAvgBytesPerSec = WaveFormat.Format.nSamplesPerSec * WaveFormat.Format.nBlockAlign;
	WaveFormat.Format.cbSize = 22;
	WaveFormat.dwChannelMask = SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT;
	WaveFormat.Samples.wValidBitsPerSample = 24;
	WaveFormat.SubFormat = KSDATAFORMAT_SUBTYPE_PCM;

	hr = RenderAudioClient->IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, (WAVEFORMATEX*)&WaveFormat, NULL);
	if (hr == AUDCLNT_E_UNSUPPORTED_FORMAT) {
		wBitsPerSample = 24;
		WaveFormat.Format.wBitsPerSample = wBitsPerSample;
		WaveFormat.Format.nBlockAlign = WaveFormat.Format.wBitsPerSample / 8 * WaveFormat.Format.nChannels;
		WaveFormat.Format.nAvgBytesPerSec = WaveFormat.Format.nSamplesPerSec * WaveFormat.Format.nBlockAlign;
	}

	//
	//  Initialize WASAPI in event driven mode, associate the audio client with our samples ready event handle, retrieve 
	//  a capture client for the transport, create the capture thread and start the audio engine.
	//
	hr = RenderAudioClient->Initialize(AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_NOPERSIST, DefaultDevicePeriod, DefaultDevicePeriod, (WAVEFORMATEX*)&WaveFormat, NULL);
	if (FAILED(hr))	{
		throw win32_error("Unable to initialize audio client", hr);
	}

	hr = RenderAudioClient->GetBufferSize(&RenderBufferSize);
	if (FAILED(hr))	{
		throw win32_error("Unable to get buffer size", hr);
	}

	hr = RenderAudioClient->SetEventHandle(AudioSamplesReadyEvent);
	if (FAILED(hr))	{
		throw win32_error("Unable to set ready event", hr);
	}

	hr = RenderAudioClient->GetService(IID_PPV_ARGS(&RenderClient));
	if (FAILED(hr)) {
		throw win32_error("Unable to get new render client", hr);
	}

	// 録音デバイス準備
	if (tr8_index_in != UINT_MAX && index != tr8_index_out) {
		init_capture(0, tr8_index_in);

	}
	if (tb3_index_in != UINT_MAX && index != tb3_index_out) {
		init_capture(1, tb3_index_in);

	}

	// 録音開始
	for (int i = 0; i < 2; i++) {
		if (CaptureAudioClient[i]) {
			hr = CaptureAudioClient[i]->Start();
			if (FAILED(hr))	{
				throw win32_error("Unable to start capture client", hr);
			}
		}
	}

	rec_data();


	// 再生開始
	RenderThread = CreateThread(NULL, 0, WASAPIRenderThread, NULL, 0, NULL);
	if (RenderThread == NULL) {
		throw win32_error("Unable to create thread", GetLastError());
	}

	// バッファを埋める
	UINT32 padding;
	hr = RenderAudioClient->GetCurrentPadding(&padding);
	if (FAILED(hr))	{
		throw win32_error("Unable to get padding", hr);
	}

	UINT32 framesAvailable = RenderBufferSize - padding;
	BYTE *pData;
	hr = RenderClient->GetBuffer(framesAvailable, (BYTE**)&pData);
	if (FAILED(hr))	{
		throw win32_error("Unable to get buffer", hr);
	}

	hr = RenderClient->ReleaseBuffer(framesAvailable, AUDCLNT_BUFFERFLAGS_SILENT);
	if (FAILED(hr))	{
		throw win32_error("Unable to release buffer", hr);
	}

	//
	//  We're ready to go, start rendering!
	//
	hr = RenderAudioClient->Start();
	if (FAILED(hr))	{
		throw win32_error("Unable to start render client", hr);
	}
}

void stop()
{
	HRESULT hr;

	// 停止
	SetEvent(ShutdownEvent);
	if (RenderAudioClient) {
		hr = RenderAudioClient->Stop();
		if (FAILED(hr))	{
			throw win32_error("Unable to stop render client", hr);
		}
	}

	if (RenderThread) {
		WaitForSingleObject(RenderThread, INFINITE);
		CloseHandle(RenderThread);
	}

	// 録音停止
	for (int i = 0; i < 2; i++) {
		if (CaptureAudioClient[i]) {
			CaptureAudioClient[i]->Stop();
		}
	}
}

void rec_data()
{
	for (int n = 0; n < 2; n++) {
		if (CaptureClient[n]) {
			HRESULT hr;
			UINT32 framesAvailable;
			BYTE *pData;
			DWORD  flags;
			while (rec_play_diff[n] < (int)RenderBufferSize) {
				hr = CaptureClient[n]->GetBuffer(&pData, &framesAvailable, &flags, NULL, NULL);
				if (FAILED(hr))	{
					continue;
				}

				for (UINT i = 0; i < framesAvailable; i++) {
					// L
					int l;
					memcpy(&l, pData + i * 2 * 4, 4);
					ring_buf[n][rec_pos[n] * 2] = l;

					// R
					int r;
					memcpy(&r, pData + i * 2 * 4 + 4, 4);
					ring_buf[n][rec_pos[n] * 2 + 1] = r;

					rec_pos[n]++;
					if (rec_pos[n] == RING_BUF_SIZE) {
						rec_pos[n] = 0;
					}
				}

				CaptureClient[n]->ReleaseBuffer(framesAvailable);

				rec_play_diff[n] += framesAvailable;
			}
		}
	}
}

void play_data()
{
	HRESULT hr;

	BYTE *pData;
	hr = RenderClient->GetBuffer(RenderBufferSize, (BYTE**)&pData);
	if (FAILED(hr))	{
		return;
	}

	for (UINT i = 0; i < RenderBufferSize; i++) {
		int l = ring_buf[0][play_pos * 2] + ring_buf[1][play_pos * 2];
		int r = ring_buf[0][play_pos * 2 + 1] + ring_buf[1][play_pos * 2 + 1];

		if (wBitsPerSample == 24) {
			// L
			BYTE *pl = (BYTE*)&l;
			pData[i * 2 * 3 + 0] = pl[1];
			pData[i * 2 * 3 + 1] = pl[2];
			pData[i * 2 * 3 + 2] = pl[3];
			// R
			BYTE *pr = (BYTE*)&r;
			pData[i * 2 * 3 + 0 + 3] = pr[1];
			pData[i * 2 * 3 + 1 + 3] = pr[2];
			pData[i * 2 * 3 + 2 + 3] = pr[3];
		}
		else {
			// L
			int *pl = (int*)(pData + i * 2 * 4);
			*pl = l;
			// R
			int *pr = (int*)(pData + i * 2 * 4 + 4);
			*pr = r;
		}


		play_pos++;
		if (play_pos == RING_BUF_SIZE) {
			play_pos = 0;
		}
	}

	RenderClient->ReleaseBuffer(RenderBufferSize, 0);

	rec_play_diff[0] -= RenderBufferSize;
	rec_play_diff[1] -= RenderBufferSize;
}

bool RenderBuffer()
{
	// 録音
	rec_data();

	// 再生
	play_data();

	return true;
}

DWORD WINAPI WASAPIRenderThread(LPVOID Context)
{
	bool stillPlaying = true;
	HANDLE waitArray[2] = { ShutdownEvent, AudioSamplesReadyEvent };
	HANDLE mmcssHandle = NULL;
	DWORD mmcssTaskIndex = 0;

	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr))	{
		return hr;
	}

	mmcssHandle = AvSetMmThreadCharacteristics(L"Audio", &mmcssTaskIndex);
	if (mmcssHandle == NULL) {
		return FALSE;
	}

	while (stillPlaying)
	{
		DWORD waitResult = WaitForMultipleObjects(2, waitArray, FALSE, INFINITE);
		switch (waitResult)
		{
		case WAIT_OBJECT_0 + 0:     // ShutdownEvent
			stillPlaying = false;
			break;
		case WAIT_OBJECT_0 + 1:     // AudioSamplesReadyEvent
			bool ret = RenderBuffer();
			if (!ret)
			{
				stillPlaying = false;
			}
			break;
		}
	}

	AvRevertMmThreadCharacteristics(mmcssHandle);
	CoUninitialize();

	return 0;
}

INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg) {
	case WM_INITDIALOG:
	{
		// 出力デバイスを追加
		for (wstring name : devices) {
			SendMessage(GetDlgItem(hwndDlg, IDC_CMB_DEVICE), CB_ADDSTRING, 0, (LPARAM)name.c_str());
		}

		// ステータス表示
		wstring str = L"TR-8 input device:\t";
		str += (tr8_index_in != UINT_MAX) ? L"OK" : L"NG";
		str += L"\r\nTB-3 input device:\t";
		str += (tb3_index_in != UINT_MAX) ? L"OK" : L"NG";
		str += L"\r\nTR-8 midi input device:\t";
		str += (tr8_midi_id_in != UINT_MAX) ? L"OK" : L"NG";
		str += L"\r\nTB-3 midi input device:\t";
		str += (tb3_midi_id_in != UINT_MAX) ? L"OK" : L"NG";
		str += L"\r\nTR-8 midi output device:\t";
		str += (tr8_midi_id_out != UINT_MAX) ? L"OK" : L"NG";
		str += L"\r\nTB-3 midi output device:\t";
		str += (tb3_midi_id_out != UINT_MAX) ? L"OK" : L"NG";

		SetWindowText(GetDlgItem(hwndDlg, IDC_STATUS), str.c_str());

		return TRUE;
	}
	case WM_COMMAND:
		switch (HIWORD(wParam)) {
		case BN_CLICKED:
			switch (LOWORD(wParam)) {
			case IDC_BTN_READY:
			{
				EnableWindow(GetDlgItem(hwndDlg, IDC_BTN_READY), FALSE);
				UINT index = (UINT)SendMessage(GetDlgItem(hwndDlg, IDC_CMB_DEVICE), CB_GETCURSEL, 0, 0);
				ready(index);
				return TRUE;
			}
			}
		}
		break;
	case WM_CLOSE:
		DestroyWindow(hwndDlg);
		return TRUE;
	case WM_DESTROY:
		PostQuitMessage(0);
		return TRUE;
	}
	return FALSE;
}

//
//  Retrieves the device friendly name for a particular device in a device collection.  
//
//  The returned string was allocated using malloc() so it should be freed using free();
//
bool GetDeviceName(IMMDeviceCollection *DeviceCollection, UINT DeviceIndex, LPWSTR deviceName, size_t size)
{
	IMMDevice *device;
	HRESULT hr;

	hr = DeviceCollection->Item(DeviceIndex, &device);
	if (FAILED(hr))	{
		return false;
	}

	IPropertyStore *propertyStore;
	hr = device->OpenPropertyStore(STGM_READ, &propertyStore);
	SafeRelease(&device);
	if (FAILED(hr))	{
		return false;
	}

	PROPVARIANT friendlyName;
	PropVariantInit(&friendlyName);
	hr = propertyStore->GetValue(PKEY_Device_FriendlyName, &friendlyName);
	SafeRelease(&propertyStore);
	if (FAILED(hr))	{
		return false;
	}

	hr = StringCbPrintf(deviceName, size, L"%s", friendlyName.vt != VT_LPWSTR ? L"Unknown" : friendlyName.pwszVal);
	if (FAILED(hr))	{
		return false;
	}

	PropVariantClear(&friendlyName);

	return true;
}

void CALLBACK MidiInProc(HMIDIIN hMidiIn, UINT wMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2)
{
	switch (wMsg)
	{
	case MIM_DATA:
	{
		BYTE status = (BYTE)dwParam1;

		if (status == 0xf8 || status == 0xfa || status == 0xfc)
		{
			// MIDIクロック
			midiOutShortMsg(hmo, status);
		}
		break;
	}
	}
}
