#include <Windows.h>
#include <stdio.h>
#include <tchar.h>

#define  PIPE_TIMEOUT 500
#define BUFSIZE 4096

typedef struct
{
    OVERLAPPED oOverlap;
    HANDLE hPipeInst;
    TCHAR chRequest[BUFSIZE];
    DWORD cbRead;
    TCHAR chReply[BUFSIZE];
    DWORD cbToWrite;
}PIPEINST, *LPPIPEINST;


void DisconnectAndClose(LPPIPEINST);
BOOL CreateAndConnectInstance(LPOVERLAPPED);
BOOL ConnectToNewClient(HANDLE, LPOVERLAPPED);
void GetAnswerToRequest(LPPIPEINST);
void WINAPI CompletedWriteRoutine(DWORD, DWORD, LPOVERLAPPED);
void WINAPI CompleteReadRoutine(DWORD, DWORD, LPOVERLAPPED);

HANDLE hPipe;


/*
����ˣ������ܵ� -> ���� ->��д ->�ر�
*/
int main()
{
    HANDLE hConnectEvent;
    OVERLAPPED oConnect;
    LPPIPEINST lpPipeInst;
    DWORD dwWait, cnRet;
    BOOL fSuccess, fPentingIo;

    //�������Ӳ������¼�����
    hConnectEvent = CreateEvent(
        NULL,
        TRUE,
        TRUE,
        NULL);

    if(hConnectEvent == NULL)
    {
        printf("createEvent failed with %d.\n",GetLastError());
        return 0;
    }
    //overlapped�¼�
    oConnect.hEvent = hConnectEvent;

    //��������ʵ�����ȴ�����
    fPentingIo = CreateAndConnectInstance(&oConnect);

    while(true)
    {
        dwWait = WaitForSingleObjectEx(hConnectEvent,INFINITE, TRUE);

        //         switch(dwWait)
        //         {
        //         case 0:
        if(dwWait == 0)
        {
            if(fPentingIo)
            {
                //��ȡoverlapped�� i / o���
                fSuccess = GetOverlappedResult(
                    hPipe,      //pipe���
                    &oConnect,  //overlapped�ṹ
                    &cnRet,     //�Ѿ����͵�������
                    FALSE);     //���ȴ�

                if(!fSuccess)
                {
                    printf("ConnnectNamePile(%d) \n.",GetLastError);
                    return 0;
                }
            }

            //�����ڴ�
            lpPipeInst = (LPPIPEINST)HeapAlloc(GetProcessHeap(), 0 , sizeof(PIPEINST));
            if(lpPipeInst == NULL)
            {
                printf("GlobalAlloc failed &d \n",GetLastError());
                return 0;
            }

            lpPipeInst->hPipeInst = hPipe;

            //����д,��д����֮����໥����
            lpPipeInst->cbToWrite = 0;
            CompletedWriteRoutine(0, 0, (LPOVERLAPPED) lpPipeInst);

            //�ٴ���һ������ʵ��������Ӧ��һ���ͻ��˵�����
            //fPentingIo = CreateAndConnectInstance(&oConnect);
           // break;
        }
        //         case WAIT_IO_COMPLETION:
        //             break;
        else if(dwWait == WAIT_IO_COMPLETION)
            break;
        else
        { 
            {
                printf("WairForSingleObjectEx %d", GetLastError());
                return 0;
            }
        }
    }

    return 0;
}


/*
д��pipe��������ɺ���
��д�������ʱ����ʼ������һ���ͻ��˵�����
*/

//�Ȳ���ͬ��i/oģ�ͽ��ܺ�д����
void WINAPI CompletedWriteRoutine(DWORD dwErr, 
    DWORD cbWritten, 
    LPOVERLAPPED lpOverlap)
{
    LPPIPEINST lpPipeInst;
    BOOL fRead = FALSE;
    //����overlapʵ��
    lpPipeInst = (LPPIPEINST)lpOverlap;

    DWORD cbRread;
    //
    if((dwErr == 0) && (cbWritten == lpPipeInst->cbToWrite))
    {
//         fRead = ReadFileEx(
//             lpPipeInst->hPipeInst,
//             lpPipeInst->chRequest,
//             BUFSIZE * sizeof(TCHAR),
//             (LPOVERLAPPED)lpPipeInst,
//             (LPOVERLAPPED_COMPLETION_ROUTINE)CompleteReadRoutine);

        //����ͬ��i/o�ܻ�ȡ������
        fRead = ReadFile(
                lpPipeInst->hPipeInst,
                lpPipeInst->chRequest,
                BUFSIZE * sizeof(TCHAR),
                &cbRread,
                NULL);

        DWORD errorWord = GetLastError();

        TCHAR *request = lpPipeInst->chRequest;
        TCHAR *reply = lpPipeInst->chReply;
        _tprintf(TEXT("%s\n"), lpPipeInst->chRequest);
        if(!fRead)
        {
            DisconnectAndClose(lpPipeInst);
        }
        else
        {
            //��ȡ�ɹ�
            BOOL bSuccess = FALSE;
            LPTSTR lpvMessage = TEXT("MESSAGE FROM server");

            bSuccess = WriteFile(
                lpPipeInst->hPipeInst,
                lpvMessage,
                ((lstrlen(lpvMessage) +1) * sizeof(TCHAR)),
                &cbWritten,
                NULL);

            if(!bSuccess)
            {
                printf("server return message falied");
               // return 
            }
        }
    }
}

/*
��ȡpipe��������ɺ���
�����������ʱ�����ã�д��ظ�
*/

void WINAPI CompleteReadRoutine(DWORD dwErr, 
    DWORD cbBytesRead, 
    LPOVERLAPPED lpOverlap)
{
    LPPIPEINST lpPipeInst;
    BOOL fWrite = FALSE;

    lpPipeInst = (LPPIPEINST) lpOverlap;

    if((dwErr == 0) && cbBytesRead != 0)
    {
        //���ݿͻ����������ɻظ�
        GetAnswerToRequest(lpPipeInst);
        fWrite = WriteFileEx(
            lpPipeInst->hPipeInst,
            lpPipeInst->chReply,
            lpPipeInst->cbToWrite,
            (LPOVERLAPPED) lpPipeInst,
            (LPOVERLAPPED_COMPLETION_ROUTINE)CompletedWriteRoutine);
    }

    if(!fWrite)
    {
        //
        DisconnectAndClose(lpPipeInst);
    }
}

/*
�Ͽ�һ�����ӵ�ʵ��

*/
void DisconnectAndClose(LPPIPEINST lpPipeInst)
{
    if(!DisconnectNamedPipe(lpPipeInst->hPipeInst))
    {
        printf("DisconnectNamedPipe failed with %d.\n", GetLastError());
    }

    CloseHandle(lpPipeInst->hPipeInst);

    if(lpPipeInst != NULL)
        HeapFree(GetProcessHeap(), 0, lpPipeInst);
}

//��������ʵ��

BOOL CreateAndConnectInstance(LPOVERLAPPED lpOverlap)
{
    LPWSTR lpszPipeName = TEXT("\\\\.\\pipe\\spmaplenamepipe");

    //���������ܵ�
    hPipe = CreateNamedPipe(
        lpszPipeName,               //pipe ��
        PIPE_ACCESS_DUPLEX |         //�ɶ���дģʽ
        FILE_FLAG_OVERLAPPED,       //overlappedģʽ
        PIPE_TYPE_MESSAGE |         //��Ϣ���� pipe
        PIPE_READMODE_MESSAGE |     //��Ϣ��ģʽ
        PIPE_WAIT,                  //����ģʽ
        PIPE_UNLIMITED_INSTANCES,   //������ʵ��
        BUFSIZE * sizeof(TCHAR),    //��������С
        BUFSIZE * sizeof(TCHAR),    //���뻺���С
        PIPE_TIMEOUT,               //�ͻ��˳�ʱ
        NULL                        //Ĭ�ϰ�ȫ����
        );

    if(hPipe == INVALID_HANDLE_VALUE)
    {
        printf("CreateNamedPipe failed with %d.\n", GetLastError()); 
        return 0;
    }

    return ConnectToNewClient(hPipe, lpOverlap);
}

//��������ʵ��

BOOL ConnectToNewClient(HANDLE hPipe, LPOVERLAPPED lpo)
{
    BOOL fConnect, fPendingIo = FALSE;

    fConnect = ConnectNamedPipe(hPipe, lpo);

    if(fConnect)
    {
        printf("ConnectNamedPipe failed with %d.\n", GetLastError()); 
        return 0;
    }

    switch(GetLastError())
    {
        // overlapped���ӽ�����. 
    case ERROR_IO_PENDING:
        fPendingIo = TRUE;
        break;
        // �Ѿ����ӣ����Eventδ��λ 
    case ERROR_PIPE_CONNECTED:
        if(SetEvent(lpo->hEvent))
            break;
    default:
        {
            printf("ConnectNamedPipe failed with %d.\n", GetLastError());
            return 0;
        }
    }

    return fPendingIo;
}

void GetAnswerToRequest(LPPIPEINST pipe)
{
    _tprintf(TEXT("[%d] %s\n", pipe->hPipeInst, pipe->chRequest));
    lstrcpyn(pipe->chReply, TEXT("DEFAULT ANSWER FROM SERVER"), BUFSIZE);
    pipe->cbToWrite = (lstrlen(pipe->chReply + 1) * sizeof(TCHAR));
}


