//�������ܵ�����ȡ��� ->д������ ->�ȴ��ظ�

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

        //�������pipe_busy����ֱ���˳�
        if(GetLastError() != ERROR_PIPE_BUSY)
        {
            printf("Could not open pipe"); 
            return 0;
           // break;
        }

        //������е�pipe�����ڷ�æ״̬���ȴ�2s
        if(!WaitNamedPipe(lpszPipeName, 2000))
        {
            printf("could not open pipe");
            return 0;
           // break;
        }
    }


    //pipe �Ѿ����ӣ�����Ϊ��Ϣ��״̬
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
        //���ظ�
        fSuccess = ReadFile(
            hPipe,
            chBuf,
            BUFSIZE * sizeof(TCHAR),
            &cbRread,
            NULL);

        if(! fSuccess && GetLastError() != ERROR_MORE_DATA)
            break;

        _tprintf( TEXT("%s\n"), chBuf ); // ��ӡ���Ľ��
    } while (!fSuccess);

    getch();

    CloseHandle(hPipe);
    system("puase");
    return 0;
}
