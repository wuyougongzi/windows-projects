//打开命名管道，获取句柄 ->写入数据 ->等待回复

#include <windows.h>
#include <stdio.h>
#include <conio.h>
#include <tchar.h>

#define  BUFSIZE 512

int main()
{
    HANDLE hPipe;
    LPTSTR lpvMessage = TEXT("MESSAGE FROM CLIENT");
    TCHAR chBuf[BUFSIZE];
    BOOL fSuccess;
    DWORD cbRread, cbWritten, dwMode;
    LPTSTR lpszPipeName = TEXT("\\\\.\\pipe\\spmaplenamepipe");

    while(1)
    {
        hPipe = CreateFile(
            lpszPipeName,
            GENERIC_READ | GENERIC_WRITE,
            0,
            NULL,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL/* | FILE_FLAG_OVERLAPPED*/,
            NULL
            );

        if(hPipe != INVALID_HANDLE_VALUE)
            break;

        //如果不是pipe_busy错误，直接退出
        if(GetLastError() != ERROR_PIPE_BUSY)
        {
            printf("Could not open pipe"); 
            return 0;
           // break;
        }

        //如果所有的pipe都处于繁忙状态，等待2s
        if(!WaitNamedPipe(lpszPipeName, 2000))
        {
            printf("could not open pipe");
            return 0;
           // break;
        }
    }


    //pipe 已经连接，设置为消息读状态
    dwMode = PIPE_READMODE_MESSAGE;
    fSuccess = SetNamedPipeHandleState(
        hPipe,
        &dwMode,
        NULL,
        NULL);

    if(!fSuccess)
    {
        printf("SetNamedPipeHandleState failed"); 
        return 0;
    }

    fSuccess = WriteFile(
        hPipe,
        lpvMessage,
        ((lstrlen(lpvMessage) +1) * sizeof(TCHAR)),
        &cbWritten,
        NULL);

//        fSuccess = WriteFileEx(
//            hPipe,
//            lpvMessage,
//            ((lstrlen(lpvMessage) +1) * sizeof(TCHAR)),
//            hPipe,
//            NULL);


    if(!fSuccess)
    {
        printf("WriteFile failed"); 
        return 0;
    }

    do 
    {
        //读回复
        fSuccess = ReadFile(
            hPipe,
            chBuf,
            BUFSIZE * sizeof(TCHAR),
            &cbRread,
            NULL);

        if(! fSuccess && GetLastError() != ERROR_MORE_DATA)
            break;

        _tprintf( TEXT("%s\n"), chBuf ); // 打印读的结果
    } while (!fSuccess);

    getch();

    CloseHandle(hPipe);
    system("puase");
    return 0;
}
