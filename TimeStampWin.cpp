///////////////////////////////////////////////////////////////////////////////////
//
// TimeStampWin.cpp 
//
// Author Peter J Lawrence Dec. 2006 Email P.J.Lawrence@gre.ac.uk
//
// This software is provided 'as-is', without any express or implied
// warranty.  In no event will the authors be held liable for any damages
// arising from the use of this software.
//
// Permission is granted to anyone to use this software for any purpose
// and to alter it and redistribute it freely, subject to the following restrictions:
//
// 1. The origin of this software must not be misrepresented; you must not
//    claim that you wrote the original software. If you use this software
//    in a product, an acknowledgment in the product documentation would be
//    appreciated but is not required.
// 2. Altered source versions must be plainly marked as such, and must not be
//    misrepresented as being the original software.
// 3. This notice may not be removed or altered from any source distribution.
//
// Modification History
//
// v1.0.0 first release (3/01/07)
// Extended Registry information to include size of window. (PJL 4/1/07)
// The "Time Control" group box added to the main dialogue window. (PJL 4/1/07)
// Changes to the open file dialogue box (removed read only option) (PJL 5/1/07)
// Fixed small bug in AddDays when working out whether a date was before or after another date
//
// v1.0.1 second release  (10/01/07)
//
// Changed Radio buttons to check boxes (PJL 23/01/07)
// Changed code to output only changed attributes (PJL 23/01/07) 
// v1.0.2 third release  (29/01/07)
//
// Included extra data in registry  (PJL 05/02/07)
// Improved text control text output (PJL 05/02/07)
// Fixed bug where application locked if it can't access a given file (PJL 05/02/07)
// Added abort functionality (PJL 06/02/07)
// Changed undo code to output only changed attributes (PJL 06/02/07)
// Fixed MAJOR BUG with current time and undo functionality assigning accessed and creation file times wrongly
// v1.0.3 release  (06/02/07)
//
// Fixed bug in the abort operation of undo
// v1.0.4 release (12/02/07)
//
// Fixed problem with Iconized Control Dialogue being closed and saving of registry (19/02/07)
// v1.0.5 release 20/02/07 
//
// Code changed for compiling under MS x64 compilers (2/03/07)
// v1.0.5 Win64 release 02/03/07
//
// Fixed problem with file time and daylight saving (23/08/07)
// v1.0.6 release 23/08/07
// 
///////////////////////////////////////////////////////////////////////////////////
// Windows Header Files:
#include <windows.h>
#include <tchar.h>

#include "resource.h"
#include "stdio.h"
#include "math.h"
#include "commctrl.h"
#include "commdlg.h"
#include <sys/stat.h>
#include <time.h>
#include "cderr.h"
#include "direct.h"
///////////////////////////////////////////////////////////////////////////////////
#include <string>
#include <vector>
#include "dlgtxtctrl.h"
///////////////////////////////////////////////////////////////////////////////////
#define MinXValDlg 366
#define MinYValDlg 200
///////////////////////////////////////////////////////////////////////////////////
// Foward declarations of functions included in this code module:
LRESULT CALLBACK	DateControlDlg(HWND, UINT, WPARAM, LPARAM);
void EnableEditBoxes(HWND hDlg, BOOL Status);
///////////////////////////////////////////////////////////////////////////////////
int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
    DialogBox(hInstance, (LPCTSTR)IDD_DATE_DIALOG, NULL, (DLGPROC)DateControlDlg);
	return 0;
}
//////////////////////////////////////////////////////////////////////////////////
BOOL SaveCurrentStatus(HWND hWnd,const std::string aWorkingFolder)
{

    RECT TheRect;
    GetWindowRect(hWnd,&TheRect);  

    LONG            lRetVal;
    HKEY            hkResult;
    lRetVal = RegCreateKeyEx( HKEY_CURRENT_USER, 
                        _T("SOFTWARE\\WinTimeStamp"), 0, NULL, 0,
                        KEY_ALL_ACCESS, NULL, &hkResult, NULL );

    if( lRetVal == ERROR_SUCCESS )
    {
        if (!IsIconic(hWnd) && !IsZoomed(hWnd))
        {
            TheRect.right-=TheRect.left;
            TheRect.bottom-=TheRect.top;
            RegSetValueEx( hkResult, _T("WinLeftPos"),NULL, REG_DWORD, (PBYTE) &TheRect.left, sizeof(TheRect.left) );
            RegSetValueEx( hkResult, _T("WinTopPos"), NULL, REG_DWORD, (PBYTE) &TheRect.top, sizeof(TheRect.top) );
            RegSetValueEx( hkResult, _T("WinWidth"),  NULL, REG_DWORD, (PBYTE) &TheRect.right, sizeof(TheRect.right) );
            RegSetValueEx( hkResult, _T("WinHeight"), NULL, REG_DWORD, (PBYTE) &TheRect.bottom, sizeof(TheRect.bottom) );
        }

        if (aWorkingFolder.length()>0)
        {
            RegSetValueEx( hkResult, _T("WorkingFolder"), NULL, REG_EXPAND_SZ, (PBYTE)aWorkingFolder.c_str(), (DWORD) aWorkingFolder.length()+1 );
        }

        LRESULT aState=SendDlgItemMessage( hWnd, IDC_Accessed,BM_GETSTATE,0,0);
        BYTE IsChk=(BYTE) 0;
        if ((aState & BST_CHECKED)) IsChk=(BYTE) 1;
        RegSetValueEx( hkResult, _T("AccessedChk"),NULL, REG_BINARY, (PBYTE) &IsChk, sizeof(IsChk) );

        aState=SendDlgItemMessage( hWnd, IDC_Modified,BM_GETSTATE,0,0);
        IsChk=(BYTE) 0;
        if ((aState & BST_CHECKED)) IsChk=(BYTE) 1;
        RegSetValueEx( hkResult, _T("ModifiedChk"),NULL, REG_BINARY, (PBYTE) &IsChk, sizeof(IsChk) );

        aState=SendDlgItemMessage( hWnd, IDC_Creation,BM_GETSTATE,0,0);
        IsChk=(BYTE) 0;
        if ((aState & BST_CHECKED)) IsChk=(BYTE) 1;
        RegSetValueEx( hkResult, _T("CreationChk"),NULL, REG_BINARY, (PBYTE) &IsChk, sizeof(IsChk) );

        aState=SendDlgItemMessage( hWnd, IDC_CurrentTime,BM_GETSTATE,0,0);
        IsChk=(BYTE) 0;
        if ((aState & BST_CHECKED)) IsChk=(BYTE) 1;
        RegSetValueEx( hkResult, _T("CurrentTimeChk"),NULL, REG_BINARY, (PBYTE) &IsChk, sizeof(IsChk) );

        RegCloseKey( hkResult );
        return(TRUE);
    }
    return(FALSE);
}
//////////////////////////////////////////////////////////////////////////////////
BOOL LoadCurrentStatus(HWND hWnd,bool RetoreWindowStatus,std::string &aWorkingFolder)
{
    LONG            lRetVal;
    HKEY            hkResult;
    DWORD size;

    lRetVal = RegCreateKeyEx( HKEY_CURRENT_USER, 
                        _T("SOFTWARE\\WinTimeStamp"), 0, NULL, 0,
                        KEY_ALL_ACCESS, NULL, &hkResult, NULL );
    if( lRetVal == ERROR_SUCCESS )
    {
        RECT TheRect;
        GetWindowRect(hWnd,&TheRect);
        TheRect.right-=TheRect.left;
        TheRect.bottom-=TheRect.top;

        size = sizeof(TheRect.left);
        lRetVal = RegQueryValueEx( hkResult, _T("WinLeftPos"),NULL, NULL, (PBYTE) &TheRect.left, &size );
        if (lRetVal == ERROR_SUCCESS)
            lRetVal = RegQueryValueEx( hkResult, _T("WinTopPos"), NULL, NULL, (PBYTE) &TheRect.top,  &size);
        if (lRetVal == ERROR_SUCCESS)
            lRetVal = RegQueryValueEx( hkResult, _T("WinWidth"),  NULL, NULL, (PBYTE) &TheRect.right, &size );
        if (lRetVal == ERROR_SUCCESS)
            lRetVal = RegQueryValueEx( hkResult, _T("WinHeight"), NULL, NULL, (PBYTE) &TheRect.bottom, &size );

        if (RetoreWindowStatus && lRetVal == ERROR_SUCCESS)
        {
            if (TheRect.left<0)
            {
                TheRect.left=0;
            }
            if (TheRect.top<0)
            {
                TheRect.top=0;
            }
            if (TheRect.right<MinXValDlg)
            {
                TheRect.right=MinXValDlg;
            }
            if (TheRect.bottom<MinYValDlg)
            {
                TheRect.bottom=MinYValDlg;
            }
            SetWindowPos(hWnd,HWND_TOP,TheRect.left,TheRect.top,TheRect.right,TheRect.bottom,SWP_SHOWWINDOW);
        }

        DWORD MaxSubKeyLen,ValuesCount,MaxValueNameLen,MaxValueLen;
        lRetVal= RegQueryInfoKey(hkResult, NULL, NULL, NULL, NULL, &MaxSubKeyLen, NULL, 
                           &ValuesCount, &MaxValueNameLen, &MaxValueLen, NULL, NULL);

        if( lRetVal == ERROR_SUCCESS )
        {
            if (MaxValueLen>0)
            {
                MaxValueLen++;
                char *DataStr=new char[MaxValueLen];
                if (DataStr)
                {
                    lRetVal = RegQueryValueEx( hkResult, _T("WorkingFolder"), NULL, NULL, (PBYTE) DataStr, &MaxValueLen );
                    if (lRetVal == ERROR_SUCCESS)
                    {
                        aWorkingFolder.assign(DataStr);
                    }
                    delete [] DataStr;
                }
            }
        }

        BYTE IsChk=(BYTE) 0;
        size = sizeof(IsChk);
        lRetVal= RegQueryValueEx(hkResult, _T("AccessedChk"),NULL, NULL, (PBYTE) &IsChk, &size );
        if( lRetVal == ERROR_SUCCESS )
        {
            if (IsChk==0)
            {
                SendDlgItemMessage( hWnd, IDC_Accessed,BM_SETCHECK,FALSE,NULL);
            }
            else
            {
                SendDlgItemMessage( hWnd, IDC_Accessed,BM_SETCHECK,TRUE,NULL);
            }
        }

        lRetVal= RegQueryValueEx(hkResult, _T("ModifiedChk"),NULL, NULL, (PBYTE) &IsChk, &size );
        if( lRetVal == ERROR_SUCCESS )
        {
            if (IsChk==0)
            {
                SendDlgItemMessage( hWnd, IDC_Modified,BM_SETCHECK,FALSE,NULL);
            }
            else
            {
                SendDlgItemMessage( hWnd, IDC_Modified,BM_SETCHECK,TRUE,NULL);
            }
        }

        lRetVal= RegQueryValueEx(hkResult, _T("CreationChk"),NULL, NULL, (PBYTE) &IsChk, &size );
        if( lRetVal == ERROR_SUCCESS )
        {
            if (IsChk==0)
            {
                SendDlgItemMessage( hWnd, IDC_Creation,BM_SETCHECK,FALSE,NULL);
            }
            else
            {
                SendDlgItemMessage( hWnd, IDC_Creation,BM_SETCHECK,TRUE,NULL);
            }
        }

        lRetVal= RegQueryValueEx(hkResult, _T("CurrentTimeChk"),NULL, NULL, (PBYTE) &IsChk, &size );
        if( lRetVal == ERROR_SUCCESS )
        {
            if (IsChk==0)
            {
                SendDlgItemMessage( hWnd, IDC_CurrentTime,BM_SETCHECK,FALSE,NULL);
                EnableEditBoxes(hWnd,TRUE);
            }
            else
            {
                SendDlgItemMessage( hWnd, IDC_CurrentTime,BM_SETCHECK,TRUE,NULL);
                EnableEditBoxes(hWnd,FALSE);
            }
        }

        RegCloseKey( hkResult );
        return(TRUE);
    }
    return (FALSE);
}
//////////////////////////////////////////////////////////////////////////////////
// options
#define OPT_MODIFY_ATIME       (1 << 0)
#define OPT_MODIFY_MTIME       (1 << 1)
#define OPT_MODIFY_CTIME       (1 << 2)
#define OPT_NO_CREATE	       (1 << 3)
//////////////////////////////////////////////////////////////////////////////////
//
// The function that does all the file attribute updating
//
// This function comes from the Touch for Windows code developed by Joergen Sigvardsson.
//
static DWORD touch(LPCTSTR lpszFile, FILETIME* atime, FILETIME* mtime, FILETIME* ctime, WORD wOpts)
{
	SetLastError(ERROR_SUCCESS);

	HANDLE hFile = CreateFile(
		lpszFile, 
		GENERIC_WRITE, 
		FILE_SHARE_READ | FILE_SHARE_WRITE, 
		NULL, 
		(wOpts & OPT_NO_CREATE) ? OPEN_EXISTING : OPEN_ALWAYS, 
		FILE_ATTRIBUTE_NORMAL, 
		0);

	DWORD dwRetVal = GetLastError();

	// Check for CreateFile() special cases
	if(dwRetVal != ERROR_SUCCESS)
    {
		if((wOpts & OPT_NO_CREATE) && dwRetVal == ERROR_FILE_NOT_FOUND)
        {
			return ERROR_SUCCESS; // not an error
        }
		else if(hFile != INVALID_HANDLE_VALUE && dwRetVal == ERROR_ALREADY_EXISTS)
        {
			dwRetVal = ERROR_SUCCESS; // not an error according to MSDN docs
        }
	}

	if(dwRetVal != ERROR_SUCCESS)
    {
		return dwRetVal;
    }

	// Is there any template timestamp?  
	if(atime || mtime || ctime)
    {
		SetLastError(ERROR_SUCCESS);
		SetFileTime(
			hFile, 
			(wOpts & OPT_MODIFY_CTIME) ? ctime : NULL,
			(wOpts & OPT_MODIFY_ATIME) ? atime : NULL,
			(wOpts & OPT_MODIFY_MTIME) ? mtime : NULL);
		dwRetVal = GetLastError();
	}

	CloseHandle(hFile);

	return dwRetVal;
}
/////////////////////////////////////////////////////////////////////////////
// Mesage handler for about box.
LRESULT CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_INITDIALOG:
				return TRUE;

		case WM_COMMAND:
			if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
			{
				EndDialog(hDlg, LOWORD(wParam));
				return TRUE;
			}
			break;
	}
    return FALSE;
}
//////////////////////////////////////////////////////////////////////////////////
//
//  Calculate modules the MS Excel way 
//  MOD(n, d) = n - d*INT(n/d)
//
int DateModules(int aNum, int aModVal)
{
    return ( (int) (aNum- aModVal*(int) floor( (double) aNum/ (double) aModVal) ) );
}
//////////////////////////////////////////////////////////////////////////////////
//
// Calculates whether the year specified is a leap year
//
bool isLeapYear(const int aYear)
{
    if ( (aYear % 4) == 0) // Almost every year which is divisible by 4 is a leap year.
    {
        if ( (aYear % 100) == 0) 
        {
            // however, may not be since almost ever year that is divisible by 100 is not
            if ( (aYear % 400) == 0)
            {
                // in this case it is a leap year since also divisible by 400,
                return (true);
            }
            return (false);
        }
        return (true);
    }
    return (false);
}
//////////////////////////////////////////////////////////////////////////////////
bool AddDays(int &TheDay, int &TheMonth, int &TheYear, int DaysToAdd, int MonthsToAdd,int YearsToAdd,int &TotalDaysAdded)
{
    int OrgDay=TheDay;
    int OrgMonth=TheMonth;
    int OrgYear=TheYear;

    TotalDaysAdded=0;

    if (TheMonth<0 || TheMonth > 12)
    {
        // not a valid month
        return (false);
    }
    const int MonthDays[12] = { 31,28,31,30,31,30,31,31,30,31,30,31};

    // Update Months
    int xYear=0;
    TheMonth+=MonthsToAdd;

    if (TheMonth<1)
    {
        xYear = TheMonth / 12 - 1;
        TheMonth = TheMonth - 1; // 0 skipped
        TheMonth = DateModules(TheMonth,12)+1; // range 1 to 12
    }
    else
    {
        xYear = (TheMonth-1) / 12; // range 1 to 12
        TheMonth = (TheMonth-1) % 12 + 1;
    }

    // Add any extra year
    YearsToAdd+=xYear;

    // Update Years
    TheYear += YearsToAdd;

    // Correct day of the month for the current year
    int DaysInMonth = MonthDays[TheMonth-1];
    if (TheMonth==2) // Feb
    {
        if (isLeapYear(TheYear))
        {
            DaysInMonth++;
        }
    }
    if (TheDay>DaysInMonth)
    {
        TheDay=DaysInMonth;
    }

    // Add Days
    int IncDay=1;
    if (DaysToAdd<0)
    {
        IncDay=-1;
    }

    int aDay;
    for (aDay=0;aDay<abs(DaysToAdd);aDay++)
    {
        TheDay+=IncDay;

        DaysInMonth = MonthDays[TheMonth-1];
        if (TheMonth==2) // Feb
        {
            if (isLeapYear(TheYear))
            {
                DaysInMonth++;
            }
        }
        if (TheDay<1 || TheDay > DaysInMonth)
        {
            if (IncDay<0)
            {
                TheMonth--;
                if (TheMonth<1)
                {
                    TheMonth=12;
                    TheYear--;
                }
                DaysInMonth = MonthDays[TheMonth-1];
                if (TheMonth==2) // Feb
                {
                    if (isLeapYear(TheYear))
                    {
                        DaysInMonth++;
                    }
                }
                TheDay=DaysInMonth;
            }
            else
            {
                TheMonth++;
                if (TheMonth>12)
                {
                    TheMonth=1;
                    TheYear++;
                }
                TheDay=1;
            }
        }
    }

    // now calculate total number of days added/removed
    IncDay=1;
    if (OrgYear>TheYear)
    {
        IncDay=-1;
    }
    else if (OrgYear==TheYear)
    {
        if (OrgMonth>TheMonth)
        {
            IncDay=-1;
        }
        else if (OrgMonth==TheMonth)
        {
            if (OrgDay>TheDay)
            {
                IncDay=-1;
            }
        }
    }

    // now count the days between  original and updated date

    TotalDaysAdded=0;
    while (OrgYear!=TheYear || OrgMonth!=TheMonth || OrgDay!=TheDay)
    {
        OrgDay+=IncDay;
        TotalDaysAdded+=IncDay;

        if (IncDay<0)
        {
            if (OrgYear<TheYear)
            {
                // problem found with date (Abort!)
                return (false);
            }
        }
        else
        {
            if (OrgYear>TheYear)
            {
                // problem found with date (Abort!)
                return (false);
            }
        }

        DaysInMonth = MonthDays[OrgMonth-1];
        if (OrgMonth==2) // Feb
        {
            if (isLeapYear(OrgYear))
            {
                DaysInMonth++;
            }
        }
        if (OrgDay<1 || OrgDay > DaysInMonth)
        {
            if (IncDay<0)
            {
                OrgMonth--;
                if (OrgMonth<1)
                {
                    OrgMonth=12;
                    OrgYear--;
                }
                DaysInMonth = MonthDays[OrgMonth-1];
                if (OrgMonth==2) // Feb
                {
                    if (isLeapYear(OrgYear))
                    {
                        DaysInMonth++;
                    }
                }
                OrgDay=DaysInMonth;
            }
            else
            {
                OrgMonth++;
                if (OrgMonth>12)
                {
                    OrgMonth=1;
                    OrgYear++;
                }
                OrgDay=1;
            }
        }
    }
    return (true);
}
//////////////////////////////////////////////////////////////////////////////////
//
//  Add the specified amount of time to a given SystemTime variable
//
bool AddTime(SYSTEMTIME  &atime,
             int wMilliseconds,int wSecond,int wMinute,int wHour,int wDay, int wMonth,int wYear)
{
    // Milliseconds
    int xSecond=0;
    wMilliseconds+=atime.wMilliseconds;
    if (wMilliseconds<0)
    {
        xSecond = wMilliseconds / 1000 - 1;
        wMilliseconds = DateModules(wMilliseconds,1000);
    }
    else
    {
        xSecond=wMilliseconds / 1000; // range (0 to 999)
        wMilliseconds = wMilliseconds % 1000;
    }
    wSecond+=xSecond;
    atime.wMilliseconds=wMilliseconds;

    // Seconds
    int xMinute=0;
    wSecond+=atime.wSecond;
    if (wSecond<0)
    {
        xMinute = wSecond / 60 - 1;
        wSecond = DateModules(wSecond,60);
    }
    else
    {
        xMinute = wSecond / 60; // range (0 to 59)
        wSecond = wSecond % 60;
    }

    // ass extra minutes 
    wMinute+=xMinute;

    atime.wSecond=wSecond;

    // Minutes
    int xHour = 0;
    wMinute+=atime.wMinute;

    if (wMinute<0)
    {
         xHour = wMinute / 60 -1;
         wMinute = DateModules(wMinute,60);
    }
    else
    {
        xHour = wMinute / 60; // range (0 to 59)
        wMinute = wMinute % 60;
    }

    // add extra hours
    wHour+=xHour;

    atime.wMinute=wMinute;

    // Hours
    int xDay=0;
    wHour+=atime.wHour;
    if (wHour<0)
    {
        xDay = wHour / 24 - 1;
        wHour = DateModules(wHour,24);
    }
    else
    {
        xDay = wHour / 24; // range (0 to 23)
        wHour = wHour % 24;
    }

    // add extra days
    wDay+=xDay;

    atime.wHour=wHour;

    // Days, Month and Year

    // This is a bit problematic since number of days in each month or year is variable

    int TheDay=atime.wDay;
    int TheMonth=atime.wMonth;
    int TheYear=atime.wYear;
    int TotalDaysAdded;
    
    if (AddDays(TheDay, TheMonth, TheYear, wDay, wMonth,wYear,TotalDaysAdded))
    {
        atime.wDay=TheDay;
        atime.wMonth=TheMonth;
        atime.wYear=TheYear;

        // find new day of the week
        TotalDaysAdded+=atime.wDayOfWeek;
        TotalDaysAdded=DateModules(TotalDaysAdded,7);
        atime.wDayOfWeek = TotalDaysAdded; // range 0 to 6
        
        return (true);
    }
    return (false);
}
//////////////////////////////////////////////////////////////////////////////////
bool MakeCTime(SYSTEMTIME &atime,char *TimeStr,int TimeStrSize)
{
    const char *Months[12] = { "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"};
    const char *Days[7] = { "Sun","Mon","Tue","Wed","Thu","Fri","Sat" };

    if (atime.wDayOfWeek<7 && atime.wMonth>0 && atime.wMonth<13)
    {
        _snprintf(TimeStr,TimeStrSize,"%03s %03s %02i %02i:%02i:%02i %04i",
            Days[atime.wDayOfWeek],Months[atime.wMonth-1],atime.wDay,atime.wHour,atime.wMinute,atime.wSecond,atime.wYear);
        return (true);
    }
    return (false);
}
/////////////////////////////////////////////////////////////////////////////
//
// Gets file times from a file
//
static bool GetFileTimes(LPCTSTR    lpszTemplateFileName,
						 FILETIME*  tsAtime,
						 FILETIME*  tsMtime,
						 FILETIME*  tsCtime) 
{
	HANDLE hFile = CreateFile(
		lpszTemplateFileName, 
		GENERIC_READ, 
		FILE_SHARE_READ | FILE_SHARE_WRITE, 
		NULL, 
		OPEN_EXISTING, 
		FILE_ATTRIBUTE_NORMAL, 
		0);

	if(INVALID_HANDLE_VALUE == hFile)
    {
		//PrintError(lpszTemplateFileName, GetLastError());
		//ShowUsage();
		return false;
	}

	if(!GetFileTime(hFile, tsCtime, tsAtime, tsMtime))
    {
		//PrintError(lpszTemplateFileName, GetLastError());
		//ShowUsage();
		CloseHandle(hFile);
		return false;
	}

	CloseHandle(hFile);

	return true;
}
//////////////////////////////////////////////////////////////////////////////////
//
//  Given a filename with a directory and extension this routine removes
//  the directory from the filename.
//  The directory is returned in the string directory.
//
bool ExtractFilenameAndDirectory(char *filename, char *directory)
{
    if (!filename) return false;
    const char DirectorySty='\\';
    char* TheLoc=strrchr(filename,DirectorySty);
    if (TheLoc>0)
    {
        size_t result = (size_t) (TheLoc - filename + 1);
        size_t i;
        for (i=0;i<result;i++)
        {
            directory[i]=filename[i];
        }
        directory[result]='\0';
        size_t j=0;
        for (i=result;i<(int) strlen(filename);i++)
        {
            filename[j]=filename[i];
            j++;
        }
        filename[j]='\0';
    }
    return(true);
}
//////////////////////////////////////////////////////////////////////////////////
void winBeep()
{
    MessageBeep(MB_ICONEXCLAMATION);
}
//////////////////////////////////////////////////////////////////////////////////
//
// this routine test to see if the input string is a number
//
// if isFloat then a floating point number is required
//
void check_for_number(HWND TheControl, BOOL isFloat, BOOL CanBeNeg)
{
    int i;

    #define SZ_NAME         20
    char buf[BUFSIZ];
    
    GetWindowText(TheControl,buf,BUFSIZ);
    if ((int)strlen(buf) > SZ_NAME)
    {
        buf[SZ_NAME] = '\0';
        SetWindowText(TheControl,buf);
        winBeep();
        MessageBox(TheControl,"Field too long","Error",MB_OK|MB_ICONINFORMATION);
    }
    buf[SZ_NAME] = '\0';
    int dp,spaces;
    dp=1;spaces=0;
    if (isFloat) dp=0;

    int MoveBack=0;
    for (i = 0; buf[i] != '\0'; i++)
    {
        if (MoveBack>0)
        {
            buf[i-MoveBack]=buf[i];
        }
        if ( !(buf[i] >= '0' && buf[i] <='9') )
        {
            if (!(CanBeNeg && spaces == 0 && buf[i] == '-'))
                if (buf[i]!=' ') spaces=1;
                
            if (!(buf[i] == '.' && dp==0) && 
                !(buf[i] == '-' && spaces==0 && CanBeNeg) &&
                !(buf[i] == ' ' && spaces==0) )
            {
                // remove this character
                MoveBack++;
            }
            if (buf[i]=='.')  dp=1; 
            if (CanBeNeg && spaces == 0 && buf[i] == '-') spaces=1;              
        } else spaces=1;
    }
    if (MoveBack>0)
    {
        // update control if charactors have been removed
        buf[i-MoveBack]='\0';
        SetWindowText(TheControl,buf);
        SetFocus(TheControl);
        SendMessage(TheControl, EM_SETSEL,i, i);
    }
}
//////////////////////////////////////////////////////////////////////////////////
#define LEFTSTICKY         0xFF000000UL
#define TOPSTICKY          0x00FF0000UL
#define RIGHTSTICKY        0x0000FF00UL
#define BOTTOMSTICKY       0x000000FFUL
#define SAMESTICKY         0xF0F0F0F0UL
#define ALLSTICKY          0xFFFFFFFFUL
#define TOPLEFTSTICKY      0xFFFF0000UL
#define TOPRIGHTSTICKY     0x00FFFF00UL
#define BOTTOMLEFTSTICKY   0xFF0000FFUL
#define BOTTOMRIGHTSTICKY  0x0000FFFFUL
#define NULLSTICKY         0x00000000UL
//////////////////////////////////////////////////////////////////////////////////
void MoveDlgItem(HWND DlgWin,unsigned long TheID,int XOffset,int YOffset, unsigned int GlueType)
{
    RECT TheRect;
    HWND aWin=GetDlgItem(DlgWin,TheID);
    GetWindowRect(aWin,&TheRect);
    POINT TheLocation;

    TheLocation.x=TheRect.left;
    TheLocation.y=TheRect.top;
    ScreenToClient(DlgWin,&TheLocation);
    TheRect.left=TheLocation.x;
    TheRect.top=TheLocation.y;
    TheLocation.x=TheRect.right;
    TheLocation.y=TheRect.bottom;
    ScreenToClient(DlgWin,&TheLocation);
    TheRect.right=TheLocation.x;
    TheRect.bottom=TheLocation.y;

    if (GlueType==TOPLEFTSTICKY)
    {
        TheRect.bottom=TheRect.bottom+YOffset;
    }
    else if (GlueType==BOTTOMRIGHTSTICKY)
    {
        TheRect.right=TheRect.right+XOffset;
        TheRect.top=TheRect.top+YOffset;
        TheRect.bottom=TheRect.bottom+YOffset;
    }
    else if (GlueType==ALLSTICKY)
    {
        TheRect.right=TheRect.right+XOffset;
        TheRect.bottom=TheRect.bottom+YOffset;
    }
    else
    {
        if ((GlueType & RIGHTSTICKY)==RIGHTSTICKY)
        {
            TheRect.right=TheRect.right+XOffset;
            TheRect.left=TheRect.left+XOffset;
        }
        if ((GlueType & BOTTOMSTICKY)==BOTTOMSTICKY)
        {
            TheRect.top=TheRect.top+YOffset;
            TheRect.bottom=TheRect.bottom+YOffset;
        }
        if ((GlueType & LEFTSTICKY)==LEFTSTICKY)
        {
        }
        if ((GlueType & TOPSTICKY)==TOPSTICKY)
        {
        }
    }
    MoveWindow(aWin,TheRect.left,TheRect.top,TheRect.right-TheRect.left,TheRect.bottom-TheRect.top,TRUE);
}
//////////////////////////////////////////////////////////////////////////////////
bool HandleSize(HWND DlgWin,int theWidth, int theHeight, int &oldWidth, int &oldHeight)
{ 
    if (IsIconic(DlgWin)) return(true);

    RECT ClientRect;
    GetClientRect(DlgWin,&ClientRect);
  
    // difference in size
    int XOffset=theWidth-oldWidth;
    int YOffset=theHeight-oldHeight;

    // ratios
    float ExpandX=theWidth/float(oldWidth);
    float ExpandY=theHeight/float(oldHeight);

    MoveDlgItem(DlgWin,IDC_EDITBOX,XOffset,YOffset,ALLSTICKY);

    MoveDlgItem(DlgWin,IDCANCEL,XOffset,YOffset,BOTTOMSTICKY);
    MoveDlgItem(DlgWin,IDUNDOABORT,XOffset,YOffset,BOTTOMSTICKY);
    MoveDlgItem(DlgWin,IDOK,XOffset,YOffset,BOTTOMSTICKY);

    oldWidth=ClientRect.right;
    oldHeight=ClientRect.bottom;
 
    InvalidateRect( DlgWin, NULL, FALSE );
    return(true);
}
//////////////////////////////////////////////////////////////////////////////////
//
// Class used to store undo information
//
class FileDateClass
{
private:
    std::string m_FileName;
    FILETIME m_tsAtime;
    FILETIME m_tsMtime;
    FILETIME m_tsCtime;
    WORD m_wOpts;

public:

    FileDateClass() { m_wOpts=0;m_FileName.assign("NULL"); }

    FileDateClass(const FileDateClass &SrcClass)
    {
        m_FileName.assign(SrcClass.m_FileName);
        m_wOpts=SrcClass.m_wOpts;
        m_tsAtime=SrcClass.m_tsAtime;
        m_tsMtime=SrcClass.m_tsMtime;
        m_tsCtime=SrcClass.m_tsCtime;
    }

    FileDateClass(const char* aFileName,WORD a_wOpts,FILETIME &a_tsAtime,FILETIME &a_tsMtime,FILETIME &a_tsCtime)
    {
        m_FileName.assign(aFileName);
        m_wOpts=a_wOpts;
        m_tsAtime=a_tsAtime;
        m_tsMtime=a_tsMtime;
        m_tsCtime=a_tsCtime;
    }

    bool GetData(char* aFileName,int FileNameSize,FILETIME &a_tsAtime,FILETIME &a_tsMtime,FILETIME &a_tsCtime)
    {
        a_tsAtime=m_tsAtime;
        a_tsMtime=m_tsMtime;
        a_tsCtime=m_tsCtime;

        strncpy(aFileName,m_FileName.c_str(),FileNameSize);
        return (true);
    }

    bool GetFileName(std::string &aFileName)
    {
        aFileName=m_FileName;
        return (true);
    }

    bool GetFileName(char* aFileName,int aFileNameSize)
    {
        strncpy(aFileName,m_FileName.c_str(),aFileNameSize);
        return (true);
    }

    bool GetFileTimes(FILETIME &a_tsAtime,FILETIME &a_tsMtime,FILETIME &a_tsCtime)
    {
        a_tsAtime=m_tsAtime;
        a_tsMtime=m_tsMtime;
        a_tsCtime=m_tsCtime;
        return (true);
    }

    WORD GetwOpts() { return (m_wOpts); }
};
//////////////////////////////////////////////////////////////////////////////////
void ProcessPendingEvents()
{
    MSG msg;         // address of structure with message
    if (PeekMessage (&msg, NULL, 0, 0,PM_REMOVE))
    {
        TranslateMessage (&msg);
        DispatchMessage (&msg);
    }
}
//////////////////////////////////////////////////////////////////////////////////
void EnableEditBoxes(HWND hDlg, BOOL Status)
{
    EnableWindow(GetDlgItem(hDlg,IDC_HoursEdit), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_MinsEdit), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_SecsEdit), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_DaysEdit), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_MonthsEdit), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_YearsEdit), Status);

    EnableWindow(GetDlgItem(hDlg,IDC_SPIN_HOURS), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_SPIN_MINS), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_SPIN_SECS), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_SPIN_DAYS), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_SPIN_MONTHS), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_SPIN_YEARS), Status);
}
//////////////////////////////////////////////////////////////////////////////////
void EnableInputControls(HWND hDlg, BOOL Status)
{
    if (Status)
    {
        LRESULT aState=SendDlgItemMessage( hDlg, IDC_CurrentTime,BM_GETSTATE,0,0);
        if ( (aState & BST_CHECKED))
        {
            EnableEditBoxes(hDlg,FALSE);
        }
        else
        {
            EnableEditBoxes(hDlg,TRUE);
        }
    }
    else
    {
        EnableEditBoxes(hDlg,FALSE);

    }
    EnableWindow(GetDlgItem(hDlg,IDC_Accessed), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_Modified), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_Creation), Status);
    EnableWindow(GetDlgItem(hDlg,IDC_CurrentTime), Status);

    EnableWindow(GetDlgItem(hDlg,IDOK), Status);
    EnableWindow(GetDlgItem(hDlg,IDCANCEL), Status);
    if (Status)
    {
        SetWindowText(GetDlgItem(hDlg,IDUNDOABORT),"Undo");
    }
    else
    {
        SetWindowText(GetDlgItem(hDlg,IDUNDOABORT),"Abort");
        EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), TRUE);
    }
}
//////////////////////////////////////////////////////////////////////////////////
//
// Mesage handler for about box.
//
LRESULT CALLBACK DateControlDlg(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    int aVal;
    static int oldWidth=0,oldHeight=0;
    static std::vector<FileDateClass> Undolist;
    static std::string WorkingFolder;
    static DlgTxtControl TxtWindow(hDlg,IDC_EDITBOX);
    static bool ContinueOperation=false;

	switch (message)
	{
		case WM_INITDIALOG:
            // set date for controls
            
            SendDlgItemMessage(hDlg, IDC_HoursEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) "0");
            SendMessage(GetDlgItem(hDlg,IDC_SPIN_HOURS), UDM_SETRANGE, 0, MAKELPARAM(48, -48));

            SendDlgItemMessage(hDlg, IDC_MinsEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) "0");
            SendMessage(GetDlgItem(hDlg,IDC_SPIN_MINS), UDM_SETRANGE, 0, MAKELPARAM(120, -120));

            SendDlgItemMessage(hDlg, IDC_SecsEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) "0");
            SendMessage(GetDlgItem(hDlg,IDC_SPIN_SECS), UDM_SETRANGE, 0, MAKELPARAM(120, -120));

            SendDlgItemMessage(hDlg, IDC_DaysEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) "0");
            SendMessage(GetDlgItem(hDlg,IDC_SPIN_DAYS), UDM_SETRANGE, 0, MAKELPARAM(50, -50));

            SendDlgItemMessage(hDlg, IDC_MonthsEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) "0");
            SendMessage(GetDlgItem(hDlg,IDC_SPIN_MONTHS), UDM_SETRANGE, 0, MAKELPARAM(24, -24));

            SendDlgItemMessage(hDlg, IDC_YearsEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) "0");
            SendMessage(GetDlgItem(hDlg,IDC_SPIN_YEARS), UDM_SETRANGE, 0, MAKELPARAM(9999, -9999));

            TxtWindow.Clear();
            TxtWindow.AppendText(_T("Status Window....\r\n"));
#ifdef _WIN64
            TxtWindow.AppendText(_T("Version: 1.0.6 Aug 2007 (Freeware) Win64\r\n"));
#else
            TxtWindow.AppendText(_T("Version: 1.0.6 Aug 2007 (Freeware) Win32\r\n"));
#endif
            TxtWindow.AppendText(_T("\r\n"));
            TxtWindow.AppendText(_T("Author: Peter J.Lawrence\r\n"));
            TxtWindow.AppendText(_T("email P.J.Lawrence@gre.ac.uk\r\n\r\n"));
            TxtWindow.AppendText(_T("This software is provided \"as is\", without any guarantee made as to its suitability or fitness for any particular use.\r\nThe Author takes no responsibility for any damage that may unintentionally be caused through its use.\r\n"));

            SendDlgItemMessage( hDlg, IDC_Modified,BM_SETCHECK,TRUE,NULL);

            EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), FALSE);

            {
                #ifdef _WIN64
                  HINSTANCE hInstance=(HINSTANCE) GetWindowLongPtr(hDlg, GWLP_HINSTANCE );
                #else
                  HINSTANCE hInstance=(HINSTANCE) GetWindowLong(hDlg, GWL_HINSTANCE );
                #endif
                HICON hIcon = LoadIcon (hInstance, MAKEINTRESOURCE(IDI_TIMESTAMPWIN));
                HICON hIconSmall = LoadIcon (hInstance, MAKEINTRESOURCE(IDI_SMALL));
                PostMessage (hDlg, WM_SETICON, ICON_BIG, (LPARAM) hIcon);
                PostMessage (hDlg, WM_SETICON, ICON_SMALL, (LPARAM) hIconSmall);
                DestroyIcon(hIcon);
                DestroyIcon(hIconSmall);

                RECT ClientRect;
                GetClientRect(hDlg,&ClientRect);
                oldWidth=ClientRect.right;
                oldHeight=ClientRect.bottom;
            }

            if (LoadCurrentStatus(hDlg,true,WorkingFolder))
            {
                if (!WorkingFolder.empty())
                {
                    if (_chdir(WorkingFolder.c_str())!=0)
                    {
                        WorkingFolder.erase();
                    }
                }
                if (WorkingFolder.empty())
                {
                    // assign working folder
                    char *Buffer;
                    if ( (Buffer=_getcwd(NULL,0)) )
                    {
                        WorkingFolder.assign(Buffer);
                        free(Buffer);
                    } 
                }
            }
            return TRUE;

        case WM_SIZING:
            {
                RECT* ClientRect = (LPRECT) lParam;
                if ( (ClientRect->right-ClientRect->left)<MinXValDlg ||
                     (ClientRect->bottom-ClientRect->top)<MinYValDlg )
                {
                    if ((ClientRect->right-ClientRect->left)<MinXValDlg)
                    {
                        ClientRect->right=ClientRect->left+MinXValDlg;
                    }
                    if ((ClientRect->bottom-ClientRect->top)<MinYValDlg)
                    {
                        ClientRect->bottom=ClientRect->top+MinYValDlg;
                    }
                    return (TRUE);
                }
            }
            break;

        case WM_SIZE:
        {
            int nWidth = LOWORD(lParam);  // width of client area 
            int nHeight = HIWORD(lParam); // height of client area
            return (HandleSize(hDlg,nWidth,nHeight,oldWidth,oldHeight));
        }

        case WM_CLOSE:
            SaveCurrentStatus(hDlg,WorkingFolder);
            EndDialog(hDlg, LOWORD(wParam));
            return (TRUE);

		case WM_COMMAND:
            switch (LOWORD(wParam))
            {
            case IDCANCEL:
                SaveCurrentStatus(hDlg,WorkingFolder);
				EndDialog(hDlg, -1);
				return TRUE;

            case IDC_Accessed:
            case IDC_Modified:
            case IDC_Creation:
                {
                    LRESULT aState=SendDlgItemMessage( hDlg, LOWORD(wParam),BM_GETSTATE,0,0);
                    if ( (aState & BST_CHECKED))
                    {
                        SendDlgItemMessage( hDlg, LOWORD(wParam),BM_SETCHECK,FALSE,NULL);
                    }
                    else
                    {
                        SendDlgItemMessage( hDlg, LOWORD(wParam),BM_SETCHECK,TRUE,NULL);
                    }
                }
                return TRUE;

            case IDC_CurrentTime:
                {
                    LRESULT aState=SendDlgItemMessage( hDlg, LOWORD(wParam),BM_GETSTATE,0,0);
                    if ( (aState & BST_CHECKED))
                    {
                        SendDlgItemMessage( hDlg, LOWORD(wParam),BM_SETCHECK,FALSE,NULL);
                        EnableEditBoxes(hDlg,TRUE);
                     }
                    else
                    {
                        SendDlgItemMessage( hDlg, LOWORD(wParam),BM_SETCHECK,TRUE,NULL);
                        EnableEditBoxes(hDlg,FALSE);
                    }
                }
                break;

            case IDC_HoursEdit:
            case IDC_MinsEdit:
            case IDC_SecsEdit:
            case IDC_DaysEdit:
            case IDC_MonthsEdit:
            case IDC_YearsEdit:
                    check_for_number((HWND) lParam,FALSE,TRUE);
                    break;

            case IDOK:
            {
                BOOL UseCurrentTime=FALSE;
                LRESULT aState=SendDlgItemMessage( hDlg, IDC_CurrentTime,BM_GETSTATE,0,0);
                if ((aState & BST_CHECKED))
                {
                    UseCurrentTime=TRUE;
                }

                WORD wOpts = (WORD) 0;
                aState=SendDlgItemMessage( hDlg, IDC_Accessed,BM_GETSTATE,0,0);
                if ((aState & BST_CHECKED))
                {
                    wOpts = wOpts | (WORD) OPT_MODIFY_ATIME;
                }
                aState=SendDlgItemMessage( hDlg, IDC_Modified,BM_GETSTATE,0,0);
                if ((aState & BST_CHECKED))
                {
				    wOpts = wOpts | (WORD) OPT_MODIFY_MTIME;
                }

                aState=SendDlgItemMessage( hDlg, IDC_Creation,BM_GETSTATE,0,0);
                if ((aState & BST_CHECKED))
                {
				    wOpts = wOpts | (WORD) OPT_MODIFY_CTIME;
                }

                #define FileNameStrSize 5000
                #define DirectoryStrSize 1000
                char TheFileName[FileNameStrSize];

                int sizeofFileName=FileNameStrSize;

                char aDirectory[DirectoryStrSize];

                ZeroMemory(TheFileName,FileNameStrSize);
                ZeroMemory(aDirectory,DirectoryStrSize);

                /* Set all structure members to zero. */
 
                const char *Filter = "All files\0";
                const char *FilterExt = "*.*\0";

                char szFileTitle[FileNameStrSize];
                #define FilterStrSize 30

                char szFilter[FilterStrSize];
                strcpy(szFilter,Filter);
                strcpy(&szFilter[strlen(Filter)+1],FilterExt);
                szFilter[strlen(szFilter)+strlen(FilterExt)+2] ='\0';
                szFileTitle[0] ='\0';

                OPENFILENAME ofn;
                memset(&ofn, 0, sizeof(OPENFILENAME));

                ofn.lStructSize = sizeof(OPENFILENAME);
                ofn.hwndOwner = GetFocus();
                ofn.lpstrFilter = szFilter;
                ofn.nFilterIndex = 1;

                ofn.lpstrFileTitle = szFileTitle;
                ofn.nMaxFileTitle = sizeof(szFileTitle);
                if (WorkingFolder.length()>0)
                {
                    ofn.lpstrInitialDir=WorkingFolder.c_str();
                }
                else
                {
                    ofn.lpstrInitialDir = aDirectory;
                }
                ofn.lpstrFile= TheFileName;
                ofn.nMaxFile = sizeofFileName;
                ofn.Flags = OFN_EXPLORER| OFN_ALLOWMULTISELECT | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
                BOOL TheStatus=GetOpenFileName(&ofn);

                if (TheStatus) 
                { 
                    TxtWindow.Clear();
                    
                    TxtWindow.AppendText(_T("Status....\r\n"));

                    size_t NextFilePos=0;
                    // test to see how many files selected
                    size_t StrLen=strlen(TheFileName);
                    if (TheFileName[StrLen+1]!=0)
                    {
                        //multiple files
                        NextFilePos=StrLen+1;
                    }

                    if (NextFilePos==0)
                    {
                        if (ExtractFilenameAndDirectory(TheFileName,aDirectory))
                        {
                            WorkingFolder.assign(aDirectory);
                        }
                        else
                        {
                            WorkingFolder.erase();
                        }
                    }
                    else
                    {
                        // first entry is Directorty others are file name
                        WorkingFolder.assign(TheFileName);

                        StrLen=strlen(&TheFileName[NextFilePos]);
                        strcpy(TheFileName,&TheFileName[NextFilePos]);

                        if (TheFileName[NextFilePos+StrLen+1]!=0)
                        {
                            // another file
                            NextFilePos=NextFilePos+StrLen+1;
                        }
                        else
                        {
                            NextFilePos=0;
                        }
                    }

                    
                    #define TheTextSize 200
                    char TheText[TheTextSize];
                    Undolist.clear();
                    EnableInputControls(hDlg,FALSE);
                    ContinueOperation=true;
                    do 
                    {
                        TxtWindow.AppendText(_T("=========================\r\n"));
                        FILETIME tsAtime, tsMtime,tsCtime;

                        if (GetFileTimes(TheFileName,&tsAtime,&tsMtime,&tsCtime) )
                        {
                            SYSTEMTIME  SysCtime;
                            FileTimeToSystemTime(&tsCtime,&SysCtime);

                            SYSTEMTIME  SysMtime;
                            FileTimeToSystemTime(&tsMtime,&SysMtime);

                            SYSTEMTIME  SysAtime;
                            FileTimeToSystemTime(&tsAtime,&SysAtime);
                            
                            #define DateStrSize 30
                            char DateStr[DateStrSize];
                            _snprintf(TheText,TheTextSize, "Accessing attributes for \"%s\"\r\n",TheFileName);
                            TxtWindow.AppendText(TheText);

                            FILETIME LocalTime;
                            SYSTEMTIME  DateTime;
                            
                            if (FileTimeToLocalFileTime(&tsAtime,&LocalTime))
                            {
                                if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                {
                                    if (MakeCTime(DateTime,DateStr,DateStrSize))
                                    {
                                        _snprintf(TheText,TheTextSize,"Accessed : %s\r\n", DateStr );
                                        TxtWindow.AppendText(TheText);
                                    }
                                }
                            }

                            if (FileTimeToLocalFileTime(&tsMtime,&LocalTime))
                            {
                                if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                {
                                    if (MakeCTime(DateTime,DateStr,DateStrSize))
                                    {
                                        _snprintf(TheText,TheTextSize,"Modified : %s\r\n", DateStr );
                                        TxtWindow.AppendText(TheText);
                                    }
                                }
                            }

                            if (FileTimeToLocalFileTime(&tsCtime,&LocalTime))
                            {
                                if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                {
                                    if (MakeCTime(DateTime,DateStr,DateStrSize))
                                    {
                                        _snprintf(TheText,TheTextSize,"Created  : %s\r\n", DateStr );
                                        TxtWindow.AppendText(TheText);
                                    }
                                }
                            }

                            if (wOpts)
                            {
                                // if we are going to change something remember current settings
                                Undolist.push_back(FileDateClass(TheFileName,wOpts,tsAtime,tsMtime,tsCtime));
                            }

                            if (UseCurrentTime)
                            {
                                if ( wOpts & OPT_MODIFY_ATIME)
                                {
                                    GetSystemTime(&SysAtime);
                                    SystemTimeToFileTime(&SysAtime,&tsAtime);
                                }

                                if ( wOpts & OPT_MODIFY_MTIME)
                                {
                                     
                                     GetSystemTime(&SysMtime);
                                     SystemTimeToFileTime(&SysMtime,&tsMtime);
                                }

                                if ( wOpts & OPT_MODIFY_CTIME)
                                {
                                     GetSystemTime(&SysCtime);
                                     SystemTimeToFileTime(&SysCtime,&tsCtime);
                                }
                            }
                            else
                            {
                                SendDlgItemMessage(hDlg, IDC_HoursEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM)TheText);
                                int hours=atoi(TheText);

                                SendDlgItemMessage(hDlg, IDC_MinsEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM)TheText);
                                int mins=atoi(TheText);

                                SendDlgItemMessage(hDlg, IDC_SecsEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM)TheText);
                                int secs=atoi(TheText);

                                SendDlgItemMessage(hDlg, IDC_DaysEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM)TheText);
                                int days=atoi(TheText);

                                SendDlgItemMessage(hDlg, IDC_MonthsEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM)TheText);
                                int months=atoi(TheText);
                            
                                SendDlgItemMessage(hDlg, IDC_YearsEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM)TheText);
                                int years=atoi(TheText);

                                if ( wOpts & OPT_MODIFY_ATIME)
                                {
                                    AddTime(SysAtime,0,secs,mins,hours,days,months,years);
                                    SystemTimeToFileTime(&SysAtime,&tsAtime);
                                }

                                if ( wOpts & OPT_MODIFY_MTIME)
                                {
                                    AddTime(SysMtime,0,secs,mins,hours,days,months,years);
                                    SystemTimeToFileTime(&SysMtime,&tsMtime);
                                }

                                if ( wOpts & OPT_MODIFY_CTIME)
                                {
                                    AddTime(SysCtime,0,secs,mins,hours,days,months,years);
                                    SystemTimeToFileTime(&SysCtime,&tsCtime);
                                }
                            }

                            DWORD dwResult = touch(
				                TheFileName,
				                (wOpts & OPT_MODIFY_ATIME) ? &tsAtime : NULL,
				                (wOpts & OPT_MODIFY_MTIME) ? &tsMtime : NULL,
				                (wOpts & OPT_MODIFY_CTIME) ? &tsCtime : NULL,
				                wOpts
			                );

			                if(ERROR_SUCCESS != dwResult) 
                            {
                                TxtWindow.AppendText(_T("FAILED to set Values\r\n"));
			                }
                            else
                            {
                                if ( (wOpts & OPT_MODIFY_ATIME) || (wOpts & OPT_MODIFY_MTIME) || (wOpts & OPT_MODIFY_CTIME))
                                {
                                    TxtWindow.AppendText(_T("Updated Values....\r\n"));

                                    if ( wOpts & OPT_MODIFY_ATIME)
                                    {
                                        if (FileTimeToLocalFileTime(&tsAtime,&LocalTime))
                                        {
                                            if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                            {
                                                if (MakeCTime(DateTime,DateStr,DateStrSize))
                                                {
                                                    _snprintf(TheText,TheTextSize,"Accessed : %s\r\n", DateStr );
                                                    TxtWindow.AppendText(TheText);
                                                }
                                            }
                                        }
                                    }

                                    if ( wOpts & OPT_MODIFY_MTIME)
                                    {
                                        if (FileTimeToLocalFileTime(&tsMtime,&LocalTime))
                                        {
                                            if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                            {
                                                if (MakeCTime(DateTime,DateStr,DateStrSize))
                                                {
                                                    sprintf(TheText,"Modified : %s\r\n", DateStr );
                                                    TxtWindow.AppendText(TheText);
                                                }
                                            }
                                        }
                                    }
                                    if ( wOpts & OPT_MODIFY_CTIME)
                                    {
                                        if (FileTimeToLocalFileTime(&tsCtime,&LocalTime))
                                        {
                                            if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                            {
                                                if (MakeCTime(DateTime,DateStr,DateStrSize))
                                                {
                                                    _snprintf(TheText,TheTextSize,"Created  : %s\r\n", DateStr );
                                                    TxtWindow.AppendText(TheText);
                                                }
                                            }
                                        } 
                                    }
                                }
                            }
                        }
                        else
                        {
                            _snprintf(TheText,TheTextSize, "FAILED to access attributes for \"%s\"\r\n",TheFileName);
                            TxtWindow.AppendText(TheText);

                        }

                        // get next file name
                        if (NextFilePos>0)
                        {
                            StrLen=strlen(&TheFileName[NextFilePos]);
                            if (StrLen>0)
                            {
                                strcpy(TheFileName,&TheFileName[NextFilePos]);
                                NextFilePos=NextFilePos+StrLen+1;
                            }
                            else
                            {
                                NextFilePos=0;
                            }
                        }

                        // Update window messages
                        ProcessPendingEvents();
                    } while (NextFilePos>0 && ContinueOperation);

                    TxtWindow.AppendText(_T("=========================\r\n"));
                    if (ContinueOperation)
                    {
                        TxtWindow.AppendText( _T("Operation Complete"));
                    }
                    else
                    {
                        TxtWindow.AppendText( _T("Aborted!"));
                    }
                }
                else
                {
                    // failed to open all files
                    DWORD WinErrorCode=CommDlgExtendedError();
                    switch (WinErrorCode)
                    {
                    case 0:
                        // Cancel pressed
                        break;
                    case FNERR_BUFFERTOOSMALL:
                        winBeep();
                        MessageBox(hDlg,"Too many files selected.\n Please select fewer files and try again.","Error",MB_OK|MB_ICONINFORMATION);
                        break;
                    default:
                        MessageBox(hDlg,"Problem selecting files.\n Please select fewer files and try again.","Error",MB_OK|MB_ICONINFORMATION);
                        break;
                    }
                }
                ContinueOperation=false;
                EnableInputControls(hDlg,TRUE);
                if (Undolist.empty())
                {
                    EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), FALSE);
                }
                else
                {
                    EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), TRUE);
                }
            }
            break;

            case IDUNDOABORT:
                if (ContinueOperation)
                {
                    // already processing so cancel current operation
                    ContinueOperation=false;
                    EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), FALSE);
                }
                else if (!Undolist.empty())
                {
                    if (MessageBox(hDlg,"Undo previous changes?","Error",MB_YESNO|MB_ICONQUESTION)==IDYES)
                    {
                        char TheText[TheTextSize];
                        EnableInputControls(hDlg,FALSE);
                        TxtWindow.Clear();
                        TxtWindow.AppendText(_T("Undo Status....\r\n"));

                        std::vector<FileDateClass>::iterator UndoIterator;
                        ContinueOperation=true;
                        unsigned int UndoCount=0;
                        for (UndoIterator =  Undolist.begin();(UndoIterator != Undolist.end() && ContinueOperation); ++UndoIterator)
                        {
                            WORD wOpts = (*UndoIterator).GetwOpts();

                            TxtWindow.AppendText(_T("=========================\r\n"));
                            FILETIME tsAtime, tsMtime,tsCtime;

                            std::string TheFileName;
                            (*UndoIterator).GetFileName(TheFileName);
                            UndoCount++;

                            if (GetFileTimes(TheFileName.c_str(),&tsAtime,&tsMtime,&tsCtime) )
                            {
                                SYSTEMTIME  SysCtime;
                                FileTimeToSystemTime(&tsCtime,&SysCtime);

                                SYSTEMTIME  SysMtime;
                                FileTimeToSystemTime(&tsMtime,&SysMtime);

                                SYSTEMTIME  SysAtime;
                                FileTimeToSystemTime(&tsAtime,&SysAtime);
                            
                                char DateStr[DateStrSize];
                                _snprintf(TheText,TheTextSize, "Accessing attributes for \"%s\"\r\n",TheFileName.c_str());
                                TxtWindow.AppendText(TheText);

                                FILETIME LocalTime;
                                SYSTEMTIME  DateTime;
                                
                                if (FileTimeToLocalFileTime(&tsAtime,&LocalTime))
                                {
                                    if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                    {
                                        if (MakeCTime(DateTime,DateStr,DateStrSize))
                                        {
                                            _snprintf(TheText,TheTextSize,"Accessed : %s\r\n", DateStr );
                                            TxtWindow.AppendText(TheText);
                                        }
                                    }
                                }

                                if (FileTimeToLocalFileTime(&tsMtime,&LocalTime))
                                {
                                    if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                    {
                                        if (MakeCTime(DateTime,DateStr,DateStrSize))
                                        {
                                            _snprintf(TheText,TheTextSize,"Modified : %s\r\n", DateStr );
                                            TxtWindow.AppendText(TheText);
                                        }
                                    }
                                }

                                if (FileTimeToLocalFileTime(&tsCtime,&LocalTime))
                                {
                                    if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                    {
                                        if (MakeCTime(DateTime,DateStr,DateStrSize))
                                        {
                                            _snprintf(TheText,TheTextSize,"Created  : %s\r\n", DateStr );
                                            TxtWindow.AppendText(TheText);
                                        }
                                    }
                                }
                            
                                (*UndoIterator).GetFileTimes(tsAtime,tsMtime,tsCtime);

                                DWORD dwResult = touch(
				                    TheFileName.c_str(),
				                    (wOpts & OPT_MODIFY_ATIME) ? &tsAtime : NULL,
				                    (wOpts & OPT_MODIFY_MTIME) ? &tsMtime : NULL,
				                    (wOpts & OPT_MODIFY_CTIME) ? &tsCtime : NULL,
				                    wOpts
			                    );

			                    if(ERROR_SUCCESS != dwResult) 
                                {
                                    TxtWindow.AppendText(_T("FAILED to set Values\r\n"));
			                    }
                                else
                                {
                                    if (wOpts>0)
                                    {
                                        if (GetFileTimes(TheFileName.c_str(),&tsAtime,&tsMtime,&tsCtime) )
                                        {
                                            TxtWindow.AppendText(_T("Updated Values (undo)....\r\n"));

                                            if ( wOpts & OPT_MODIFY_ATIME)
                                            {
                                                if (FileTimeToLocalFileTime(&tsAtime,&LocalTime))
                                                {
                                                    if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                                    {
                                                        if (MakeCTime(DateTime,DateStr,DateStrSize))
                                                        {
                                                            _snprintf(TheText,TheTextSize,"Accessed : %s\r\n", DateStr );
                                                            TxtWindow.AppendText(TheText);
                                                        }
                                                    }
                                                }
                                            }

                                            if ( wOpts & OPT_MODIFY_MTIME)
                                            {
                                                if (FileTimeToLocalFileTime(&tsMtime,&LocalTime))
                                                {
                                                    if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                                    {
                                                        if (MakeCTime(DateTime,DateStr,DateStrSize))
                                                        {
                                                            _snprintf(TheText,TheTextSize,"Modified : %s\r\n", DateStr );
                                                            TxtWindow.AppendText(TheText);
                                                        }
                                                    }
                                                }
                                            }

                                            if ( wOpts & OPT_MODIFY_CTIME)
                                            {
                                                if (FileTimeToLocalFileTime(&tsCtime,&LocalTime))
                                                {
                                                    if (FileTimeToSystemTime(&LocalTime,&DateTime))
                                                    {
                                                        if (MakeCTime(DateTime,DateStr,DateStrSize))
                                                        {
                                                            _snprintf(TheText,TheTextSize,"Created  : %s\r\n", DateStr );
                                                            TxtWindow.AppendText(TheText);
                                                        }
                                                    }
                                                } 
                                            }
                                        }
                                        else
                                        {
                                            _snprintf(TheText,TheTextSize, "FAILED to access updated attributes for \"%s\"\r\n",TheFileName.c_str());
                                            TxtWindow.AppendText(TheText);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                _snprintf(TheText,TheTextSize, "FAILED to access attributes for \"%s\"\r\n",TheFileName.c_str());
                                TxtWindow.AppendText(TheText);
                            }

                            // Update window messages
                            ProcessPendingEvents();
                        }
                        TxtWindow.AppendText(_T("=========================\r\n"));
                        if (ContinueOperation)
                        {
                            Undolist.clear();
                            TxtWindow.AppendText( _T("Undo Complete"));;
                        }
                        else
                        {
                            // remove any operations from the undo list were processed
                            if (UndoCount>0)
                            {
                                UndoIterator = Undolist.begin()+(UndoCount-1);
                                Undolist.erase(Undolist.begin(),UndoIterator);
                            }
                            TxtWindow.AppendText( _T("Undo Aborted!"));
                        }
                        EnableInputControls(hDlg,TRUE);
                        ContinueOperation=false;
                        if (Undolist.size()>0)
                        {
                            EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), TRUE);
                        }
                        else
                        {
                            EnableWindow(GetDlgItem(hDlg,IDUNDOABORT), FALSE);
                        }
                    }
                }
                break;

            case IDC_SPIN_HOURS:
                {
                    char TheText[TheTextSize];
                    SendDlgItemMessage(hDlg, IDC_HoursEdit, WM_GETTEXT, (WPARAM) TheTextSize, (LPARAM) TheText);
                    aVal=atoi(TheText);
                    aVal++;
                    _snprintf(TheText,TheTextSize,"%i",aVal);
                    SendDlgItemMessage(hDlg, IDC_HoursEdit, WM_SETTEXT, (WPARAM) FALSE, (LPARAM) TheText);
                }
                return TRUE;

            default:
                return TRUE;
			}
			break;
	}
    return FALSE;
}