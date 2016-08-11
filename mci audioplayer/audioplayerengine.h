#ifndef AUDIOPLAYERENGINE_H
#define AUDIOPLAYERENGINE_H

//对应上面的位置
struct MediaPlayerStatus 
{
	long fileLength;		//播放长度		//单位秒
	long position;			//播放位置

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
	void play();								//从头播放
	void playFrom(int position);				//从指定位置开始播放
	void puase();								//暂停
	void stop();								//停止
	bool checkCurrentStatus(MediaPlayerStatus &currentStatus);
	bool mediaFileLength(int &length);
	void audioClose();

private:
	MCI_OPEN_PARMS m_mciOpenParms;
	MCI_PLAY_PARMS m_mciPlayParms;
};
#endif //AUDIOPLAYERENGINE_H