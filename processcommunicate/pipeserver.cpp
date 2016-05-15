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
服务端：创建管道 -> 监听 ->读写 ->关闭
*/
int main()
{
    HANDLE hConnectEvent;
    OVERLAPPED oConnect;
    LPPIPEINST lpPipeInst;
    DWORD dwWait, cnRet;
    BOOL fSuccess, fPentingIo;

    //用于连接操作的事件对象
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
    //overlapped事件
    oConnect.hEvent = hConnectEvent;

    //创建连接实例，等待链接
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
                //获取overlapped的 i / o结果
                fSuccess = GetOverlappedResult(
                    hPipe,      //pipe句柄
                    &oConnect,  //overlapped结构
                    &cnRet,     //已经传送的数据量
                    FALSE);     //不等待

                if(!fSuccess)
                {
                    printf("ConnnectNamePile(%d) \n.",GetLastError);
                    return 0;
                }
            }

            //分配内存
            lpPipeInst = (LPPIPEINST)HeapAlloc(GetProcessHeap(), 0 , sizeof(PIPEINST));
            if(lpPipeInst == NULL)
            {
                printf("GlobalAlloc failed &d \n",GetLastError());
                return 0;
            }

            lpPipeInst->hPipeInst = hPipe;

            //读和写,读写函数之间的相互调用
            lpPipeInst->cbToWrite = 0;
            CompletedWriteRoutine(0, 0, (LPOVERLAPPED) lpPipeInst);

            //再创建一个连接实例，以响应下一个客户端的连接
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
写入pipe操作的完成函数
当写操作完成时，开始读另外一个客户端的请求
*/

//先采用同步i/o模型接受和写数据
void WINAPI CompletedWriteRoutine(DWORD dwErr, 
    DWORD cbWritten, 
    LPOVERLAPPED lpOverlap)
{
    LPPIPEINST lpPipeInst;
    BOOL fRead = FALSE;
    //保存overlap实例
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

        //采用同步i/o能获取到数据
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
            //读取成功
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
读取pipe操作的完成函数
当读操作完成时被调用，写入回复
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
        //根据客户端请求，生成回复
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
断开一个连接的实例

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

//建立连接实例

BOOL CreateAndConnectInstance(LPOVERLAPPED lpOverlap)
{
    LPWSTR lpszPipeName = TEXT("\\\\.\\pipe\\spmaplenamepipe");

    //创建命名管道
    hPipe = CreateNamedPipe(
        lpszPipeName,               //pipe 名
        PIPE_ACCESS_DUPLEX |         //可读可写模式
        FILE_FLAG_OVERLAPPED,       //overlapped模式
        PIPE_TYPE_MESSAGE |         //消息类型 pipe
        PIPE_READMODE_MESSAGE |     //消息读模式
        PIPE_WAIT,                  //阻塞模式
        PIPE_UNLIMITED_INSTANCES,   //无限制实例
        BUFSIZE * sizeof(TCHAR),    //输出缓存大小
        BUFSIZE * sizeof(TCHAR),    //输入缓存大小
        PIPE_TIMEOUT,               //客户端超时
        NULL                        //默认安全属性
        );

    if(hPipe == INVALID_HANDLE_VALUE)
    {
        printf("CreateNamedPipe failed with %d.\n", GetLastError()); 
        return 0;
    }

    return ConnectToNewClient(hPipe, lpOverlap);
}

//建立连接实例

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
        // overlapped连接进行中. 
    case ERROR_IO_PENDING:
        fPendingIo = TRUE;
        break;
        // 已经连接，因此Event未置位 
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


