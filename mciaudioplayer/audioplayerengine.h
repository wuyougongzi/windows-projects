#ifndef AUDIOPLAYERENGINE_H
#define AUDIOPLAYERENGINE_H

//��Ӧ�����λ��
struct MediaPlayerStatus 
{
	long fileLength;		//���ų���		//��λ��
	long position;			//����λ��

	MediaPlayerStatus()
	{
		fileLength = -1;
		position = -1;
	}
};

class AudioPlayerEngine
{
public:
	AudioPlayerEngine();
	~AudioPlayerEngine();

public:
	bool initPlayerEngine(wstring filePath);
	void play();								//��ͷ����
	void playFrom(int position);				//��ָ��λ�ÿ�ʼ����
	void puase();								//��ͣ
	void stop();								//ֹͣ
	bool checkCurrentStatus(MediaPlayerStatus &currentStatus);
	bool mediaFileLength(int &length);
	void audioClose();

private:
	MCI_OPEN_PARMS m_mciOpenParms;
	MCI_PLAY_PARMS m_mciPlayParms;
};
#endif //AUDIOPLAYERENGINE_H