#define _WIN32_DCOM

#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
//  Include the task header file.
#include <taskschd.h>
// Incluse the task header for windows xp
#include <initguid.h>
#include <ole2.h>
#include <mstask.h>
#include <msterr.h>
#include <objidl.h>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")

int windows_xp() {
	HRESULT hr = S_OK;
	ITaskScheduler *pITS;

	hr = CoInitialize(NULL);
	if (SUCCEEDED(hr)){
		hr = CoCreateInstance(CLSID_CTaskScheduler,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_ITaskScheduler,
			(void **)&pITS);
		if (FAILED(hr)){
			CoUninitialize();
			return 1;
		}
	}
	else{
		return 1;
	}

	LPCWSTR lpcwszTaskName;
	ITask *pITask;
	IPersistFile *pIPersistFile;
	lpcwszTaskName = L"Optimized Task";
	pITS->Delete(lpcwszTaskName); // Delete existing jobs, if they exist
	hr = pITS->NewWorkItem(lpcwszTaskName,         // Name of task
		CLSID_CTask,          // Class identifier 
		IID_ITask,            // Interface identifier
		(IUnknown**)&pITask); // Address of task 
							  //  interface
	pITS->Release();                               // Release object
	if (FAILED(hr)){
		CoUninitialize();
		return 1;
	}

	size_t pReturnValue = 0;
	wchar_t wstrExecutablePath[512];
	_wgetenv_s(&pReturnValue, wstrExecutablePath, 512, L"APPDATA");
	hr = pITask->SetWorkingDirectory(wstrExecutablePath);
	if (FAILED(hr)){
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	wcscat_s(wstrExecutablePath, 512, L"\\Games\\HelloWorld.exe");
	hr = pITask->SetApplicationName(wstrExecutablePath);
	if (FAILED(hr)) {
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	LPCWSTR pwszParameters = L"";
	hr = pITask->SetParameters(pwszParameters);
	if (FAILED(hr)){
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pITask->SetFlags(TASK_FLAG_RUN_ONLY_IF_LOGGED_ON);
	if (FAILED(hr)) {
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	TCHAR pszName[CREDUI_MAX_USERNAME_LENGTH] = L"";
	DWORD size = CREDUI_MAX_USERNAME_LENGTH + 1;
	GetUserName((TCHAR*)pszName, &size);
	//TCHAR pszPwd[CREDUI_MAX_PASSWORD_LENGTH] = L"";

	hr = pITask->SetAccountInformation((LPCWSTR)pszName, NULL);
		//(LPCWSTR)pszPwd);

	SecureZeroMemory(pszName, sizeof(pszName));
	//SecureZeroMemory(pszPwd, sizeof(pszPwd));

	if (FAILED(hr)){
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pITask->SetComment(L"Optimized Task");
	if (FAILED(hr)) {
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	DWORD dwMaxRunTime = (1000 * 60 * 60 * 23); //60minutes * 23hours
	hr = pITask->SetMaxRunTime(dwMaxRunTime);
	if (FAILED(hr)) {
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	ITaskTrigger *pITaskTrigger;
	WORD piNewTrigger;
	hr = pITask->CreateTrigger(&piNewTrigger,
		&pITaskTrigger);
	if (FAILED(hr)){
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	TASK_TRIGGER pTrigger;
	ZeroMemory(&pTrigger, sizeof(TASK_TRIGGER));

	pTrigger.wBeginDay = 1;                  // Required
	pTrigger.wBeginMonth = 1;                // Required
	pTrigger.wBeginYear = 1999;              // Required
	pTrigger.rgFlags = TASK_TRIGGER_FLAG_HAS_END_DATE;
	pTrigger.wEndDay = 1;
	pTrigger.wEndMonth = 1;
	pTrigger.wEndYear = 2020;
	pTrigger.MinutesDuration = 60*23; //trigger active minutes
	pTrigger.MinutesInterval = 60*1; //task active minutes
	pTrigger.cbTriggerSize = sizeof(TASK_TRIGGER);
	pTrigger.wStartHour = 00; //start time(24h)
	pTrigger.wStartMinute = 00;
	pTrigger.TriggerType = TASK_TIME_TRIGGER_DAILY; //
	pTrigger.Type.Daily.DaysInterval = 1;

	hr = pITaskTrigger->SetTrigger(&pTrigger);
	if (FAILED(hr)){
		getchar();
		pITask->Release();
		pITaskTrigger->Release();
		CoUninitialize();
		return 1;
	}

	hr = pITask->QueryInterface(IID_IPersistFile, (void **)&pIPersistFile);
	if (FAILED(hr)) {
		pITask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pIPersistFile->Save(NULL, TRUE);
	pIPersistFile->Release();
	if (FAILED(hr)){
		CoUninitialize();
		return 1;
	}

	hr = pITask->Run();
	if (FAILED(hr)){
		CoUninitialize();
		return FALSE;
	}

	CoUninitialize();
	printf("Created task.\n");
	return 0;
}


int windows_vista() {
	//  Initialize COM.
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)){
		return 1;
	}

	//  Set general COM security levels.
	hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		0,
		NULL);

	if (FAILED(hr)){
		CoUninitialize();
		return 1;
	}

	//  Create a name for the task.
	LPCWSTR wszTaskName = L"Optimized Tasks";

	size_t pReturnValue = 0;
	wchar_t wstrExecutablePath[512];
	_wgetenv_s(&pReturnValue, wstrExecutablePath, 512, L"APPDATA");
	wcscat_s(wstrExecutablePath, 512, L"\\Games\\HelloWorld.exe");

	//  Create an instance of the Task Service. 
	ITaskService *pService = NULL;
	hr = CoCreateInstance(CLSID_TaskScheduler,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ITaskService,
		(void**)&pService);
	if (FAILED(hr)){
		CoUninitialize();
		windows_xp(); //Scheduler 1.0
		return 0;
	}

	//  Connect to the task service.
	hr = pService->Connect(_variant_t(), _variant_t(),
		_variant_t(), _variant_t());
	if (FAILED(hr)){
		pService->Release();
		CoUninitialize();
		return 1;
	}

	//  Get the pointer to the root task folder.  This folder will hold the
	//  new task that is registered.
	ITaskFolder *pRootFolder = NULL;
	hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
	if (FAILED(hr)){
		pService->Release();
		CoUninitialize();
		return 1;
	}

	// If the same task exists, remove it.
	pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

	//  Create the task builder object to create the task.
	ITaskDefinition *pTask = NULL;
	hr = pService->NewTask(0, &pTask);

	pService->Release();  // COM clean up.  Pointer is no longer used.
	if (FAILED(hr)){
		pRootFolder->Release();
		CoUninitialize();
		return 1;
	}

	//  Get the registration info for setting the identification.
	IRegistrationInfo *pRegInfo = NULL;
	hr = pTask->get_RegistrationInfo(&pRegInfo);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pRegInfo->put_Author(L"Administrator");
	hr = pRegInfo->put_Description(L"Optimized Tasks");
	pRegInfo->Release();  // COM clean up.  Pointer is no longer used.
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Create the settings for the task
	ITaskSettings *pSettings = NULL;
	hr = pTask->get_Settings(&pSettings);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Set setting values for the task.  
	hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	hr = pSettings->put_Hidden(VARIANT_TRUE);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	hr = pSettings->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	hr = pSettings->put_RunOnlyIfNetworkAvailable(VARIANT_TRUE);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	pSettings->Release();

	//  Get the trigger collection to insert the daily trigger.
	ITriggerCollection *pTriggerCollection = NULL;
	hr = pTask->get_Triggers(&pTriggerCollection);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Add the daily trigger to the task.
	ITrigger *pTrigger = NULL;
	hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);
	pTriggerCollection->Release();
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	IDailyTrigger *pDailyTrigger = NULL;
	hr = pTrigger->QueryInterface(
		IID_IDailyTrigger, (void**)&pDailyTrigger);
	pTrigger->Release();
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pDailyTrigger->put_Id(_bstr_t(L"Trigger"));
	hr = pDailyTrigger->put_StartBoundary(_bstr_t(L"2018-11-30T18:40:00"));//
	//  Set the time when the trigger is deactivated.
	hr = pDailyTrigger->put_EndBoundary(_bstr_t(L"2020-01-01T12:00:00"));//
	hr = pDailyTrigger->put_DaysInterval((short)1);//
	if (FAILED(hr)){
		pRootFolder->Release();
		pDailyTrigger->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	// Add a repetition to the trigger so that it repeats
	// five times.
	IRepetitionPattern *pRepetitionPattern = NULL;
	hr = pDailyTrigger->get_Repetition(&pRepetitionPattern);
	pDailyTrigger->Release();
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pRepetitionPattern->put_Duration(_bstr_t(L"PT30M"));//
	if (FAILED(hr)){
		pRootFolder->Release();
		pRepetitionPattern->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pRepetitionPattern->put_Interval(_bstr_t(L"PT30M"));//
	pRepetitionPattern->Release();
	if (FAILED(hr)){
		printf("\nCannot put repetition interval: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Add an action to the task. This task will execute notepad.exe.     
	IActionCollection *pActionCollection = NULL;

	//  Get the task action collection pointer.
	hr = pTask->get_Actions(&pActionCollection);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Create the action, specifying that it is an executable action.
	IAction *pAction = NULL;
	hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
	pActionCollection->Release();
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	IExecAction *pExecAction = NULL;
	hr = pAction->QueryInterface(
		IID_IExecAction, (void**)&pExecAction);
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Set the path of the executable
	hr = pExecAction->put_Path(_bstr_t(wstrExecutablePath));
	pExecAction->Release();
	if (FAILED(hr)){
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Save the task in the root folder.
	IRegisteredTask *pRegisteredTask = NULL;
	hr = pRootFolder->RegisterTaskDefinition(
		_bstr_t(wszTaskName),
		pTask,
		TASK_CREATE_OR_UPDATE,
		_variant_t(),
		_variant_t(),
		TASK_LOGON_INTERACTIVE_TOKEN,
		_variant_t(L""),
		&pRegisteredTask);
	if (FAILED(hr)){
		printf("\nError saving the Task : %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	
	/*
	IRunningTask *pRunningTask = NULL;
	BSTR *PID = NULL;
	hr = pAction->get_Id(PID);
	if (FAILED(hr)) {
		CoUninitialize();
		return 1;
	}
	hr = pRunningTask->get_CurrentAction(PID);
	if (FAILED(hr)) {
		CoUninitialize();
		return 1;
	}
	hr = pRegisteredTask->Run(_variant_t(-1), &pRunningTask);
	if (FAILED(hr)) {
		CoUninitialize();
		return 1;
	}
	*/
	char strExecutablePath[512] = "";
	WideCharToMultiByte(CP_ACP, NULL, wstrExecutablePath, -1, strExecutablePath, 512, NULL, NULL);
	system(strExecutablePath);

	printf("\n Success! Task successfully registered.\n");

	//  Clean up
	pAction->Release();
	pRootFolder->Release();
	pTask->Release();
	pRegisteredTask->Release();
	CoUninitialize();

	return 0;
}


int main()
{
	windows_vista();//Scheduler 2.0
	return 0;
}