#include "audioplayerengine.h"

AudioPlayerEngine::AudioPlayerEngine()
{
	m_mciOpenParms.wDeviceID = 0;
}

AudioPlayerEngine::~AudioPlayerEngine()
{
}

bool AudioPlayerEngine::initPlayerEngine(wstring filePath)
{
	m_mciOpenParms.lpstrDeviceType = NULL;
	m_mciOpenParms.lpstrElementName = filePath.c_str();
	m_mciOpenParms.wDeviceID = 0;

	MCIERROR mciError;
	mciError = mciSendCommand(MCI_DEVTYPE_WAVEFORM_AUDIO, MCI_OPEN, MCI_OPEN_ELEMENT, (DWORD)&m_mciOpenParms);
	if(mciError != 0)
		return false;

	return true;
}

void AudioPlayerEngine::play()
{
	MCIERROR mciError;
	m_mciPlayParms.dwFrom = 0;

	mciError = mciSendCommand(m_mciOpenParms.wDeviceID, MCI_PLAY, MCI_FROM | MCI_NOTIFY, (DWORD)&m_mciPlayParms);
	if(mciError != 0)
		return;
}

void AudioPlayerEngine::playFrom(int position)
{
	MCIERROR mciError;
	m_mciPlayParms.dwFrom = position;

	mciError = mciSendCommand(m_mciOpenParms.wDeviceID, MCI_PLAY, MCI_FROM | MCI_NOTIFY, (DWORD)&m_mciPlayParms);
	if(mciError != 0)
		return;
}

void AudioPlayerEngine::puase()
{
	mciSendCommand(m_mciOpenParms.wDeviceID, MCI_PAUSE, NULL,(DWORD)&m_mciOpenParms);
}

void AudioPlayerEngine::stop()
{
	mciSendCommand(m_mciOpenParms.wDeviceID, MCI_STOP, NULL, NULL);
}

void AudioPlayerEngine::audioClose()
{
	MCIERROR mciError;
	mciError = 	mciSendCommand(m_mciOpenParms.wDeviceID, MCI_CLOSE, NULL, NULL);
}

bool AudioPlayerEngine::mediaFileLength(int &length)
{
	MCIERROR mciError;
	MCI_STATUS_PARMS statusParams;
	statusParams.dwItem = MCI_STATUS_LENGTH;
	mciError = mciSendCommand(m_mciOpenParms.wDeviceID, MCI_STATUS, MCI_WAIT | MCI_STATUS_ITEM, (DWORD)&statusParams);
	if(mciError != 0)
		return false;

	length = statusParams.dwReturn;
	return true;
}

bool AudioPlayerEngine::checkCurrentStatus(MediaPlayerStatus &currentStatus)
{
	MCIERROR mciError;
	MCI_STATUS_PARMS statusParams;

	statusParams.dwItem = MCI_STATUS_LENGTH;
	mciError = mciSendCommand(m_mciOpenParms.wDeviceID, MCI_STATUS, MCI_WAIT | MCI_STATUS_ITEM, (DWORD)&statusParams);
	if(mciError != 0)
		return false;
	currentStatus.fileLength = statusParams.dwReturn;

	statusParams.dwItem = MCI_STATUS_POSITION;
	mciError = mciSendCommand(m_mciOpenParms.wDeviceID, MCI_STATUS, MCI_WAIT | MCI_STATUS_ITEM, (DWORD)&statusParams);
	if(mciError != 0)
		return false;

	currentStatus.position = statusParams.dwReturn;

	return true;
}