//////////////////////////////
//   Wallpaper Changer
//////////////////////////////



// Some magic needed for LoadIconMetric not to throw error while linking app
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")




// COM definitions - for Wallpaper management object
#include "shobjidl.h"

// shell objects - needed for common directories in Windows
#include <Shlobj.h>

// GDI stuff for JPEG reading
#include <Gdiplus.h>
using namespace Gdiplus;

// std stuff
#include <fstream>
#include <string>
#include <map>
#include <set>
#include <vector>
using namespace std;

// Needed for logging timestamps
#include <time.h>

// Resources
#include "resource.h"




// Who we are...
const wchar_t* APP_NAME = L"Wallpaper Changer";



// Constant for converting minutes to miliseconds
#define ONE_MINUTE_MILLIS               60000
// Delay of the wallpaper update when display or HW configuration is discovered
// ...to aggregate multiple events OS fires quickly one after another
#define SET_WALLPAPER_TIMER_DELAY       3000

// ID of timer triggered after HW change
#define EVENT_SET_WALLPAPER_HW_CHANGE   0x5109
// ID of timer for periodic wallpaper updates
#define EVENT_SET_WALLPAPER_SCHEDULED   0x510A

// Message from our tray icon
#define MY_TRAY_MESSAGE                 WM_USER + 1
#define MY_MSG_FOLDER_CHANGED           WM_USER + 2

// Try icon menu IDs
#define MENU_ID_EXIT                    1
#define MENU_ID_SETTINGS                2
#define MENU_ID_SET_WALLPAPER           3



// Forward declarations for functions that do the actual job
void readWallpapers();
bool setWallpapers(bool change = 0);
unsigned long manageFolderWatcher(HWND window, bool start = true);



#pragma region "DEBUG LOGGER"

// A slow, primitive, always flushing, non thread safe file logger
class Logger
{
public:
    Logger(const WCHAR* name, bool enable = true) : enabled(enable) {
        PWSTR logDir;
        SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &logDir);
        file = logDir;
        CoTaskMemFree(logDir);
        file += L"\\";
        file += name;
        file += L".log";
    }

    const Logger& operator<< (const WCHAR* msg) const {
        if (!enabled)
            return *this;

        WCHAR buffer[80];
        formatTimestamp(buffer);

        wofstream f(file, wofstream::app);
        f << buffer << msg << L"\n";
        f.close();

        return *this;
    }

    const Logger& operator<< (const int num) const {
        if (!enabled)
            return *this;

        WCHAR buffer[80];
        formatTimestamp(buffer);

        wofstream f(file, wofstream::app);
        f << buffer << to_wstring(num).c_str() << L"\n";
        f.close();

        return *this;
    }

    void enable(bool enable = true) {
        enabled = enable;
    }

private:
    void formatTimestamp(wchar_t* buff) const {
        time_t rawtime;
        struct tm timeinfo;
        time(&rawtime);
        localtime_s(&timeinfo, &rawtime);
        wcsftime(buff, 80, L"%F %T   ", &timeinfo);
    }

    bool    enabled;
    wstring file;
} LOG(APP_NAME, true);

#pragma endregion



#pragma region "GENERIC SETTINGS UTILITIES"

// Class representing a single configurable value: integer, boolean or text.
// I think I liked the version that was storing values simply as strings better - was much much shorter.
class Value
{
public:
    typedef enum { String, Int, Bool /*, Enum*/ } Type;

    Value(bool default) : type(Type::Bool), boolVal(default) {}
    Value(int default) : type(Type::Int), intVal(default) {}
    Value(const wchar_t* default) : type(Type::String), strVal(new wstring(default)) {}
    Value() { throw std::runtime_error("never to be used");  }

    ~Value() { if (type == Type::String && strVal != nullptr) delete strVal; }

    // Get rid of the asignment operator and copy constructor
    Value & operator=(const Value&) = delete;
    Value(const Value& org) = delete;

    // Move constructor - needed since we store settings in tuple and don't want copy
    Value(Value&& org) {
        type = org.type;
        strVal = org.strVal;
        org.strVal = nullptr;
    }

public:
    operator bool() const {
        validateType(Type::Bool);
        //return value == L"true";
        return boolVal;
    }

    operator int() const {
        validateType(Type::Int);
        //return std::stoi(value);
        return intVal;
    }

    operator const wchar_t*() const {
        validateType(Type::String);
        //return value.c_str();
        return strVal->c_str();
    }

    void setVal(bool val) {
        validateType(Type::Bool);
        boolVal = val;
    }

    void setVal(int val) {
        validateType(Type::Int);
        intVal = val;
    }

    void setVal(const wchar_t* val) {
        validateType(Type::String);
        *strVal = val;
    }

    void fromString(const wchar_t* str) {
        switch (type) {
        case Type::Bool:
            boolVal = wcscmp(str, L"true") == 0;
            break;
        case Type::Int:
            intVal = std::stoi(str);
            break;
        case Type::String:
            *strVal = str;
            break;
        }
    }

    void toString(wstring& outStr) const {
        switch (type) {
        case Type::Bool:
            outStr = boolVal ? L"true" : L"false";
            break;
        case Type::Int:
            outStr = std::to_wstring(intVal);
            break;
        case Type::String:
            outStr = *strVal;
            break;
        }
    }


private:
    void validateType(Type t) const {
        if (type != t) throw std::runtime_error("incorrect type of the settings variable");
    }

    Type            type;

    union {
        bool     boolVal;
        int      intVal;
        wstring* strVal;
    };
};



template<typename T>
class Settings
{
public:
    Settings(tuple<T, const wchar_t*, Value>* tuples, int count, const wchar_t* fileName)
    {
        for (int i = 0; i < count; i++) {
            this->key2setting.insert(make_pair(std::get<0>(tuples[i]), &std::get<2>(tuples[i])));
            this->name2key.insert(make_pair(std::get<1>(tuples[i]), std::get<0>(tuples[i])));
        }

        PWSTR settingsDir;
        SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &settingsDir);
        file = settingsDir;
        CoTaskMemFree(settingsDir);
        file += L"\\";
        file += fileName;
        file += L".ini";

        load();
        save();
    }

    Value& get(T key) {
        return *key2setting[key];
    }


    void save() {
        wstring val;
        wofstream f(file);
        for (auto const& s : name2key) {
            key2setting[s.second]->toString(val);
            f << s.first << L"=" << val << std::endl;
        }
        f.close();
    }

private:
    void load() {
        wstring line;
        wstring name;
        wstring value;

        wifstream f(file);
        while (std::getline(f, line))
        {
            // Split the ini file line
            size_t pos = line.find(L"=");
            if (pos == wstring::npos) {
                // Exactly two tokens should be there
                continue;
            }
            name = line.substr(0, pos);
            value = line.substr(pos + 1);

            try {
                T key = name2key.at(name.c_str());
                Value* s = key2setting[key];
                s->fromString(value.c_str());
            }
            catch (...) {
                // Ignore setting not recognized
            }
        }
        f.close();
    }

private:
    map<T, Value*>    key2setting;
    map<wstring, T>     name2key;
    wstring             file;

};

#pragma endregion



#pragma region "APPLICATION SPECIFIC SETTINGS + CONFIG DIALOG"

typedef enum {
    ImageDirectory,
    AllowUpscaling,
    AutoChangeImage,
    AutoChangeInterval,
    AllowedAspectRatioMismatch,
    DisplayMode,
    MultiMonPolicy,
    EnableDebugLog
} WallSettings;


typedef enum {
    Different = 0,
    Same,
    Whatever
} MultiMonImage;



// Some settings that are needed for this application
Settings<WallSettings>& getSettings()
{
    PWSTR picturesDir;
    SHGetKnownFolderPath(FOLDERID_Pictures, 0, NULL, &picturesDir);

    static wstring picturesDirStr(picturesDir);
    CoTaskMemFree(picturesDir);

    static tuple<WallSettings, const wchar_t*, Value> mySettings[] = {
        make_tuple(WallSettings::ImageDirectory,             L"ImageDirectory",              Value(picturesDirStr.c_str())),
        make_tuple(WallSettings::AllowUpscaling,             L"AllowUpscaling",              Value(false)),
        make_tuple(WallSettings::AutoChangeImage,            L"AutoChangeImage",             Value(false)),
        make_tuple(WallSettings::AutoChangeInterval,         L"AutoChangeInterval",          Value(10)),
        make_tuple(WallSettings::AllowedAspectRatioMismatch, L"AllowedAspectRatioMismatch",  Value(1)),
        make_tuple(WallSettings::DisplayMode,                L"DisplayMode",                 Value(DWPOS_FILL)),
        make_tuple(WallSettings::MultiMonPolicy,             L"MultiMonPolicy",              Value(0)),
        make_tuple(WallSettings::EnableDebugLog,             L"EnableDebugLog",              Value(false)),
    };

    static Settings<WallSettings> theSettingsObj(mySettings, sizeof(mySettings) / sizeof(mySettings[0]), APP_NAME);

    return theSettingsObj;
}

Settings<WallSettings>& SETTINGS = getSettings();



void enableAutoStart(bool ena) {
    HKEY key = NULL;
    RegCreateKey(HKEY_CURRENT_USER, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", &key);
    if (ena) {
        wchar_t exePath[MAX_PATH + 1];
        GetModuleFileName(NULL, exePath, MAX_PATH + 1);
        RegSetValueEx(key, APP_NAME, 0, REG_SZ, (BYTE*)exePath, (DWORD)(wcslen(exePath) + 1) * 2);
    }
    else {
        RegDeleteValue(key, APP_NAME);
    }
}

bool isAutoStartEnabled() {
    HKEY key = NULL;
    int err = RegOpenKeyEx(HKEY_CURRENT_USER, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_READ, &key);
    if (err == ERROR_SUCCESS) {
        DWORD type;
        err = RegQueryValueEx(key, APP_NAME, NULL, &type, NULL, NULL);
        if (err == ERROR_SUCCESS && type == REG_SZ) {
            return true;
        }
    }
    return false;
}


void initDlgAndLoadSettings(HWND window)
{
    // Most settings are accessed via SETTINGS global, autostart goes directly to/from registry
    SetDlgItemText(window, IDC_DIRECTORY, SETTINGS.get(WallSettings::ImageDirectory));

    // No DWPOS_SPAN or DWPOS_TILE since the program installs separate wallpapers on every screen
    const wchar_t* modesNames[] = {
        L"Fill", // - cover entire screen, preserve aspect ratio, image may be clipped"
        L"Fit", // - preserve aspect ratio, do not clip image - black bands possible around",
        L"Stretch", // - scale to entire screen, no clipping, don't care about aspect ratio",
        L"Center", // - don't scale, center of image at center of the screen",
    };
    DESKTOP_WALLPAPER_POSITION modeCodes[] = { DWPOS_FILL, DWPOS_FIT, DWPOS_STRETCH, DWPOS_CENTER };
    HWND comboDisplayMode = GetDlgItem(window, IDC_DISPLAY_MODE);
    int sel = 0;
    for (int i = sizeof(modesNames) / sizeof(modesNames[0]) - 1; i >= 0; i--) {
        SendMessage(comboDisplayMode, CB_INSERTSTRING, 0, (LPARAM)modesNames[i]);
        SendMessage(comboDisplayMode, CB_SETITEMDATA, 0, modeCodes[i]);
        if ((int)SETTINGS.get(WallSettings::DisplayMode) == modeCodes[i]) {
            sel = i;
        }
    }
    SendMessage(comboDisplayMode, CB_SETCURSEL, sel, 0);

    CheckDlgButton(window, IDC_AUTO_START, isAutoStartEnabled() ? BST_CHECKED : BST_UNCHECKED);

    CheckDlgButton(window, IDC_AUTO_CHANGE, (bool)SETTINGS.get(WallSettings::AutoChangeImage) ? BST_CHECKED : BST_UNCHECKED);
    SendMessage(GetDlgItem(window, IDC_AUTO_CHANGE_SPIN), UDM_SETRANGE32, 0, MAXINT32);
    // Setting buddy makes the editbox too narrow
    //SendMessage(GetDlgItem(window, IDC_AUTO_CHANGE_SPIN), UDM_SETBUDDY, (int)GetDlgItem(window, IDC_AUTO_CHANGE_INTERVAL), 0);
    SetDlgItemInt(window, IDC_AUTO_CHANGE_INTERVAL, (int)SETTINGS.get(WallSettings::AutoChangeInterval), false);

    CheckDlgButton(window, IDC_ALLOW_UPSCALLING, (bool)SETTINGS.get(WallSettings::AllowUpscaling) ? BST_CHECKED : BST_UNCHECKED);

    const wchar_t* multimonNames[] = { L"different images", L"the same image", L"(no preference)" };
    MultiMonImage multiMonPolicy = MultiMonImage::Whatever;
    HWND comboMultiMon = GetDlgItem(window, IDC_MULTIPLE_MONITORS);
	for (int i = sizeof(multimonNames) / sizeof(multimonNames[0]) - 1; i >= 0;  i--) {
        SendMessage(comboMultiMon, CB_INSERTSTRING, 0, (LPARAM)multimonNames[i]);
        SendMessage(comboMultiMon, CB_SETITEMDATA, 0, multiMonPolicy);
        if ((int)SETTINGS.get(WallSettings::MultiMonPolicy) == multiMonPolicy) {
            sel = i;
        }
        ((int&)multiMonPolicy)--;
    }
    SendMessage(comboMultiMon, CB_SETCURSEL, sel, 0);

    HWND aspect = GetDlgItem(window, IDC_ASPECT_RATIO_MISMATCH);
    SendMessage(aspect, TBM_SETRANGE, true, MAKELONG(0, 1000));  // min. & max. positions
    SendMessage(aspect, TBM_SETPOS, true, (int)SETTINGS.get(WallSettings::AllowedAspectRatioMismatch));

    CheckDlgButton(window, IDC_DEBUG_LOG, (bool)SETTINGS.get(WallSettings::EnableDebugLog) ? BST_CHECKED : BST_UNCHECKED);

}


bool saveSettingsFromDlg(HWND window)
{
    wchar_t directory[MAX_PATH + 1];
    GetDlgItemText(window, IDC_DIRECTORY, directory, sizeof(directory));
    DWORD dwAttrib = GetFileAttributes(directory);
    if ((dwAttrib == INVALID_FILE_ATTRIBUTES) || (dwAttrib & FILE_ATTRIBUTE_DIRECTORY) == 0) {
        // No such directory
        MessageBox(window, L"The directory does not exist", L"Wrong value", MB_OK | MB_ICONEXCLAMATION);
        return false;
    }

    int displayPos = (int)SendMessage(GetDlgItem(window, IDC_DISPLAY_MODE), CB_GETCURSEL, 0, 0);
    int displayMode = (int)SendMessage(GetDlgItem(window, IDC_DISPLAY_MODE), CB_GETITEMDATA, displayPos, 0);

    bool autoStart = IsDlgButtonChecked(window, IDC_AUTO_START);

    bool autoChange = IsDlgButtonChecked(window, IDC_AUTO_CHANGE);
    BOOL autoChangeIntervalOk;
    int autoChangeInterval = GetDlgItemInt(window, IDC_AUTO_CHANGE_INTERVAL, &autoChangeIntervalOk, false);
    if (!autoChangeIntervalOk) {
        if (autoChange) {
            // Valid integer required, notify that we have a problem
            MessageBox(window, L"Image change interval must be a number grater than 0", L"Wrong value", MB_OK | MB_ICONEXCLAMATION);
            return false;
        }
    }

    bool allowUpscalling = IsDlgButtonChecked(window, IDC_ALLOW_UPSCALLING);

    int multiMonPos = (int)SendMessage(GetDlgItem(window, IDC_MULTIPLE_MONITORS), CB_GETCURSEL, 0, 0);
    int multiMonMode = (int)SendMessage(GetDlgItem(window, IDC_MULTIPLE_MONITORS), CB_GETITEMDATA, multiMonPos, 0);

    HWND aspect = GetDlgItem(window, IDC_ASPECT_RATIO_MISMATCH);
    int aspectMismatch = (int)SendMessage(aspect, TBM_GETPOS, 0, 0);

    bool debug = IsDlgButtonChecked(window, IDC_DEBUG_LOG);

    // All data ok, save the settings
    SETTINGS.get(WallSettings::ImageDirectory).setVal(directory);
    SETTINGS.get(WallSettings::DisplayMode).setVal(displayMode);
    enableAutoStart(autoStart);
    SETTINGS.get(WallSettings::AutoChangeImage).setVal(autoChange);
    SETTINGS.get(WallSettings::AutoChangeInterval).setVal(autoChangeIntervalOk ? autoChangeInterval : 10);
    SETTINGS.get(WallSettings::AllowUpscaling).setVal(allowUpscalling);
    SETTINGS.get(WallSettings::MultiMonPolicy).setVal(multiMonMode);
    SETTINGS.get(WallSettings::AllowedAspectRatioMismatch).setVal(aspectMismatch);
    SETTINGS.get(WallSettings::EnableDebugLog).setVal(debug);
    // In case of debug log configure the logging object
    LOG.enable(debug);

    SETTINGS.save();

    return true;
}



// Reccomended way to use OS built-in folder selection dialog (copied from MSDN).
// They don't distinguish between good and evil anymore.
bool selectDirectory(wstring& dir)
{
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        IFileOpenDialog *pFileOpen;

        // Create the FileOpenDialog object.
        hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
            IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));

        if (SUCCEEDED(hr))
        {
            DWORD dwOptions;
            if (SUCCEEDED(pFileOpen->GetOptions(&dwOptions)))
            {
                pFileOpen->SetOptions(dwOptions | FOS_PICKFOLDERS);

                // Show the Open dialog box.
                hr = pFileOpen->Show(NULL);

                // Get the file name from the dialog box.
                if (SUCCEEDED(hr))
                {
                    IShellItem *pItem;
                    hr = pFileOpen->GetResult(&pItem);
                    if (SUCCEEDED(hr))
                    {
                        PWSTR pszFilePath;
                        hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

                        // Display the file name to the user.
                        if (SUCCEEDED(hr))
                        {
                            //MessageBox(NULL, pszFilePath, L"File Path", MB_OK);
                            dir = pszFilePath;
                            CoTaskMemFree(pszFilePath);
                        }
                        pItem->Release();
                    }
                }
            }
            pFileOpen->Release();
        }
        CoUninitialize();
    }
    return SUCCEEDED(hr);
}



INT_PTR CALLBACK DialogProc(HWND window, unsigned int msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_INITDIALOG:
            initDlgAndLoadSettings(window);
            return TRUE;

        case WM_CLOSE:
            DestroyWindow(window);
            return TRUE;

        case WM_DESTROY:
            return TRUE;

        case WM_KEYDOWN:
            if (VK_CANCEL == wp) {
                EndDialog(window, IDCANCEL);
                return TRUE;
            }
            return FALSE;

        case WM_COMMAND:
            switch (LOWORD(wp))
            {
            case IDCANCEL:
                EndDialog(window, IDCANCEL);
                return TRUE;
            case IDOK:
                if (saveSettingsFromDlg(window)) {
                    EndDialog(window, IDOK);
                }
                return TRUE;
            case IDC_SELECT_DIR:
            {
                wstring dir;
                if (selectDirectory(dir)) {
                    SetDlgItemText(window, IDC_DIRECTORY, dir.c_str());
                }
            }
            }
            break;

        case WM_NOTIFY:
        {
            int code = ((LPNMHDR)lp)->code;
            switch (code)
            {
                case UDN_DELTAPOS:
                {
                    LPNMUPDOWN ud = (LPNMUPDOWN)lp;
                    BOOL autoChangeIntervalOk;
                    int autoChangeInterval = GetDlgItemInt(window, IDC_AUTO_CHANGE_INTERVAL, &autoChangeIntervalOk, false);
                    if (autoChangeIntervalOk) {
                        SetDlgItemInt(window, IDC_AUTO_CHANGE_INTERVAL, autoChangeInterval + ud->iDelta, false);
                    }
                    break;
                }
            }
        }
    }

    return FALSE;
}


void showConfig(HWND window)
{
    static bool showing = false;
    if (showing) {
    }
    else {
        // poorman's synchronization
        showing = true;
        wstring oldDir(SETTINGS.get(WallSettings::ImageDirectory));
        INT_PTR res = DialogBox(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_SETTINGS), window, DialogProc);
        if (res == IDOK) {
            readWallpapers();
            // Set the wallpaper - if directory changed force changing the image
            bool imageDirChanged = oldDir != (const wchar_t*)SETTINGS.get(WallSettings::ImageDirectory);
            setWallpapers(imageDirChanged);
            if (imageDirChanged) {
                manageFolderWatcher(window);
            }
            if ((bool)SETTINGS.get(WallSettings::AutoChangeImage)) {
                SetTimer(window, EVENT_SET_WALLPAPER_SCHEDULED, ONE_MINUTE_MILLIS * (int)SETTINGS.get(WallSettings::AutoChangeInterval), NULL);
            }
            else {
                KillTimer(window, EVENT_SET_WALLPAPER_SCHEDULED);
            }
        }
        showing = false;
    }
}

#pragma endregion



#pragma region "MONITORING FOLDER WITH IMAGES"

typedef struct {
    HANDLE          mutex;
    const wchar_t*  folder;
    HWND            window;
} WatcherData;


unsigned long WINAPI folderWatcherThreadProc(void* data)
{
    unsigned long waitStatus;
    HANDLE        changeHandle;
    WatcherData   watcherData = *(WatcherData*)data;

    changeHandle = FindFirstChangeNotification(
        watcherData.folder,                // directory to watch
        FALSE,                              // do not watch subtree
        FILE_NOTIFY_CHANGE_FILE_NAME);      // watch file name changes

    HANDLE  waitHandles[] = { changeHandle , watcherData.mutex};

    while (true)
    {
        LOG << L"Watching folder " << watcherData.folder;
        waitStatus = WaitForMultipleObjects(2, waitHandles, FALSE, INFINITE);

        switch (waitStatus)
        {
            case WAIT_OBJECT_0:
                LOG << L"Change in observed folder";
                // Notify App window about the change
                SendMessage(watcherData.window, MY_MSG_FOLDER_CHANGED, 0, 0);
                if (FindNextChangeNotification(changeHandle) == FALSE) {
                    return 0;
                }
                break;

            case WAIT_OBJECT_0 + 1:
                LOG << L"Thread asked to terminate";
                CloseHandle(watcherData.mutex);
                FindCloseChangeNotification(changeHandle);
                return 0;

            default:
                LOG << L"Unexpected notification";
                FindCloseChangeNotification(changeHandle);
                return 0;
        }
    }
}


unsigned long manageFolderWatcher(HWND window, bool start)
{
    static WatcherData data = { 0 };

    // If thread is already working ask it politely to terminate
    if (data.mutex != NULL) {
        ReleaseMutex(data.mutex);
        data.mutex = 0;
    }

    if (!start) {
        return NULL;
    }

    // The thread will not use the data anymore - use the structure again
    data.mutex = CreateMutex(NULL, TRUE, L"");
    if (data.mutex == NULL) {
        LOG << L"Creating mutex for folder watcher syncyng failed";
        return NULL;
    }

    data.window = window;
    data.folder = SETTINGS.get(WallSettings::ImageDirectory);


    unsigned long threadId;
    HANDLE thread = CreateThread(
        NULL,                       // default security attributes
        0,                          // use default stack size
        folderWatcherThreadProc,    // thread function name
        &data,                      // argument to thread function
        0,                          // use default creation flags
        &threadId);   // returns the thread identifier

    if (threadId == NULL)
    {
        LOG << L"Creating thread for watching folder changes";
    }

    LOG << L"Folder Watcher thread created";

    return threadId;
}

#pragma endregion



#pragma region "WALLPAPER IMAGES HANDLING"

// Global:
// List of known images and their dimensions
map<wstring, int> file2dimensions;

// Read all images in the configured wallpapers directory and prapare
// map of picture name to their dimentions ratio
void readWallpapers()
{
    file2dimensions.clear();

    // For reading image properties GDI library will be used - it must by initialized first
    GdiplusStartupInput gdiplusStartupInput;
    ULONG_PTR gdiplusToken;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    WIN32_FIND_DATA ffd;
    HANDLE hFind = INVALID_HANDLE_VALUE;

    const wchar_t* imageDir = SETTINGS.get(WallSettings::ImageDirectory);

    // Find the first file in the directory.
    wstring searchTemplate(imageDir);
    searchTemplate += L"\\*.jpg";
    hFind = FindFirstFile(searchTemplate.c_str(), &ffd);

    if (INVALID_HANDLE_VALUE == hFind)
    {
        // No files in this directory or no access
        return;
    }

    // List all the files in the directory with some info about them.
    do
    {
        wstring filePath(imageDir);
        filePath += L"\\";
        filePath += ffd.cFileName;
        Image* img = Image::FromFile(filePath.c_str());
        if (img == nullptr) {
            continue;
        }
        UINT dimensions = (img->GetWidth() << 16) + img->GetHeight();
        file2dimensions.insert(std::make_pair(filePath, dimensions));
        delete img;
    } while (FindNextFile(hFind, &ffd) != 0);

    // De-initialize GDI
    GdiplusShutdown(gdiplusToken);
}


// Set best wallpapers for currently attached monitors
// If change parameter is true the function will try not to use currently set wallpapers
bool setWallpapers(bool change)
{
    LOG << L"setWallpapers()";

    bool allowUpscaling = SETTINGS.get(WallSettings::AllowUpscaling);
    int allowedMismatch = SETTINGS.get(WallSettings::AllowedAspectRatioMismatch);
    MultiMonImage multiMonMode = (MultiMonImage)(int)SETTINGS.get(WallSettings::MultiMonPolicy);

    HRESULT hr;
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    IDesktopWallpaper* pWall = nullptr;
    hr = CoCreateInstance(__uuidof(DesktopWallpaper), nullptr, CLSCTX_ALL, __uuidof(IDesktopWallpaper), reinterpret_cast<LPVOID *>(&pWall));
    if (FAILED(hr) || pWall == nullptr) {
        return false;
    }

    set<const wchar_t*> used;

    UINT nMonitors = 0;
    pWall->GetMonitorDevicePathCount(&nMonitors);
    for (UINT monitor = 0; monitor < nMonitors; monitor++)
    {
        LPWSTR pId;
        hr = pWall->GetMonitorDevicePathAt(monitor, &pId);
        if (!FAILED(hr)) {
            RECT rect;
            hr = pWall->GetMonitorRECT(pId, &rect);
            if (!FAILED(hr)) {
                int ratio = 1000 * (rect.right - rect.left) / (rect.bottom - rect.top);

                bool allowUpscalling = SETTINGS.get(WallSettings::AllowUpscaling);

                set<const WCHAR*> properImages;
                for (auto const& f2d : file2dimensions)
                {
                    int w = f2d.second >> 16;
                    int h = f2d.second & 0xFFFF;

                    if (!allowUpscaling) {
                        if ((w < rect.right - rect.left) || (h < rect.bottom - rect.top)) {
                            continue;
                        }
                    }

                    int fileRatio = 1000 * w / h;
                    int mismatch = 1000 * (fileRatio - ratio) / ratio;
                    if (mismatch < 0) {
                        mismatch = -mismatch;
                    }
                    if (mismatch < allowedMismatch) {
                        properImages.insert(f2d.first.c_str());
                    }
                }

                if (properImages.size() == 1) {
                    // Only one good image found - just set it
                    const wchar_t* image = *properImages.begin();
                    hr = pWall->SetWallpaper(pId, image);
                    used.insert(image);
                }
                else if (properImages.size() > 1) {
                    // More than one matching options, choose right image depending on 'change' parameter
                    LPWSTR current;
                    pWall->GetWallpaper(pId, &current);
                    bool currentFound = false;
                    auto it = properImages.begin();
                    do {
                        if (wcscmp(current, *it) == 0) {
                            currentFound = true;
                            properImages.erase(it);
                            break;
                        }
                    } while (++it != properImages.end());
                    if (!change && currentFound) {
                        // The function was requested not to change Wallpaper and we found out, that
                        // curently set wallpaper is present in the images set - no action required
                    }
                    else {
                        // Multiple matching images available
                        if (multiMonMode == MultiMonImage::Different) {
                            // If prefference is to use different images on each
                            // screen remove from the pool already used ones
                            for (const auto & img : used) {
                                properImages.erase(img);
                            }
                            // ...but make sure something has left
                            if (properImages.size() == 0) {
                                properImages = used;
                            }
                        }
                        else if (multiMonMode == MultiMonImage::Same) {
                            // If prefference is to use the same image check if there is an intersection
                            // in sets of proper images and already used ones
                            set<const wchar_t*> usedAndProper;
                            for (const auto & img : used) {
                                if (properImages.find(img) != properImages.end()) {
                                    usedAndProper.insert(img);
                                }
                            }
                            if (usedAndProper.size() > 0) {
                                // Some of the used images are proper - so use them
                                properImages = usedAndProper;
                            }
                        }

                        // Choose random  image from available pool
                        int r = rand() % properImages.size();
                        it = properImages.begin();
                        advance(it, r);

                        hr = pWall->SetWallpaper(pId, *it);
                        used.insert(*it);
                    }
                }
                else {
                    // No suitable images for this screen - log error
                }
            }
        }
    }
    pWall->SetPosition((DESKTOP_WALLPAPER_POSITION)(int)SETTINGS.get(WallSettings::DisplayMode));

    pWall->Release();
    CoUninitialize();

    return true;
}

#pragma endregion



#pragma region "WINDOWS APPLICATION + GUI"


bool createNotificationIcon(HWND window, const WCHAR* tip)
{
    LOG << L"Creating notification icon";

    HICON icon = 0;
    LoadIconMetric(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_WALL), LIM_SMALL, &icon);

    NOTIFYICONDATA nid = {};
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.uVersion = NOTIFYICON_VERSION_4;
    // System uses window handle and ID to identify the icon - we just have one so ID == 0
    nid.hWnd = window;
    nid.uID = 0;
    nid.hIcon = icon;
    // Following message will be sent to our main window when user interacts with the tray icon
    nid.uCallbackMessage = MY_TRAY_MESSAGE;
    wcscpy_s(nid.szTip, tip);
    // Add the icon with tooltip and sending messagess to the parent window
    nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
    return Shell_NotifyIcon(NIM_ADD, &nid);

    // Set the version
    Shell_NotifyIcon(NIM_SETVERSION, &nid);
}



// Get rid of the icon in tray
bool deleteNotificationIcon(HWND window)
{
    LOG << L"Deleting notification icon";
    // System uses window handle and ID to identify the icon
    NOTIFYICONDATA nid = {};
    nid.hWnd = window;
    nid.uID = 0;
    // Delete the icon
    return Shell_NotifyIcon(NIM_DELETE, &nid);
}



// Show context menu for the tray icon
void showContextMenu(HWND window)
{
    // They say it is needed...
    SetForegroundWindow(window);

    // Point at which to display the menu - not very carelully choosen..
    POINT point;
    GetCursorPos(&point);

    HMENU menu = CreatePopupMenu();
    AppendMenu(menu, MF_STRING, MENU_ID_SET_WALLPAPER, L"Set wallpaper");
    AppendMenu(menu, MF_STRING, MENU_ID_SETTINGS, L"Settings...");
    AppendMenu(menu, MF_STRING, MENU_ID_EXIT, L"Exit");
    TrackPopupMenu(menu, TPM_CENTERALIGN | TPM_VCENTERALIGN | TPM_LEFTBUTTON, point.x, point.y, 0, window, NULL);

    DestroyMenu(menu);
}



LRESULT __stdcall WndProc(HWND window, unsigned int msg, WPARAM wp, LPARAM lp)
{
    // Windows Explorer (not Windows OS!) message - needed to handle case Explorer is restarted
    const static UINT WM_TASKBARCREATED = ::RegisterWindowMessage(L"TaskbarCreated");

    switch (msg)
    {
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;

        case WM_KILLFOCUS:
            ShowWindow(window, SW_HIDE);
            return 0;

        case WM_TIMER:
            LOG << L"WM_TIMER";
            if (wp == EVENT_SET_WALLPAPER_HW_CHANGE || wp == EVENT_SET_WALLPAPER_SCHEDULED) {
                // If the timer was caused by HW change kill it - it is one time only event
                KillTimer(window, EVENT_SET_WALLPAPER_HW_CHANGE);
                setWallpapers(wp == EVENT_SET_WALLPAPER_SCHEDULED);
            }
            return 0;

        case WM_COMMAND:
            if (HIWORD(wp) == 0) {
                switch (LOWORD(wp)) {
                    case MENU_ID_EXIT:
                        PostQuitMessage(0);
                        break;
                    case MENU_ID_SETTINGS:
                        //ShowWindow(window, SW_SHOW);
                        showConfig(window);
                        break;
                    case MENU_ID_SET_WALLPAPER:
                        setWallpapers(true);
                        // If periodic change of wallpapers is configured schedule the next update
                        if (SETTINGS.get(WallSettings::AutoChangeImage)) {
                            SetTimer(window, EVENT_SET_WALLPAPER_SCHEDULED, ONE_MINUTE_MILLIS * (int)SETTINGS.get(WallSettings::AutoChangeInterval), NULL);
                        }
                        break;
                }
            }
            return 0;

        case MY_TRAY_MESSAGE:
            switch (lp)
            {
                case WM_LBUTTONDBLCLK:
                    setWallpapers(true);
                    // If periodic change of wallpapers is configured schedule (postpone) the next update
                    if (SETTINGS.get(WallSettings::AutoChangeImage)) {
                        SetTimer(window, EVENT_SET_WALLPAPER_SCHEDULED, ONE_MINUTE_MILLIS * (int)SETTINGS.get(WallSettings::AutoChangeInterval), NULL);
                    }
                    break;
                case WM_RBUTTONDOWN:
                case WM_CONTEXTMENU:
                    showContextMenu(window);
                    break;
            }
            return 0;

        case MY_MSG_FOLDER_CHANGED:
            readWallpapers();
            setWallpapers();
            return 0;

        case WM_DISPLAYCHANGE:
        case WM_DEVICECHANGE:
            // Event that may require the wallpaper update happened - so let's do it!
            LOG << ((msg == WM_DEVICECHANGE) ?  L"WM_DEVICECHANGE" : L"WM_DISPLAYCHANGE");
            SetTimer(window, EVENT_SET_WALLPAPER_HW_CHANGE, SET_WALLPAPER_TIMER_DELAY, NULL);
            return 0;

        case WM_SETTINGCHANGE:
            if (wp == SPI_SETDESKWALLPAPER) {
                // Event that may require the wallpaper update happened - so let's do it!
                LOG << L"WM_SETTINGCHANGE ";
                SetTimer(window, EVENT_SET_WALLPAPER_HW_CHANGE, SET_WALLPAPER_TIMER_DELAY, NULL);
                return 0;
            }
    }

    // Dynamically obtained message ID - cannot use in switch above
    if (msg == WM_TASKBARCREATED) {
        // Explorer was restarted - create the icon again (just in case delete it first)
        deleteNotificationIcon(window);
        createNotificationIcon(window, L"Double-click to set desktop wallpaper");
        return 0;
    }

    return DefWindowProc(window, msg, wp, lp);
}



// Register window class and create main application window
HWND createWindow(const WCHAR* name, WNDPROC wndProc)
{
    WNDCLASSEX wndclass = { sizeof(WNDCLASSEX), CS_DBLCLKS, wndProc,
        0, 0, GetModuleHandle(0), LoadIcon(0,IDI_APPLICATION),
        LoadCursor(0,IDC_ARROW), HBRUSH(COLOR_WINDOW + 1),
        0, name, LoadIcon(0,IDI_APPLICATION) };

    if (RegisterClassEx(&wndclass))
    {
        UINT flags = WS_OVERLAPPED | WS_CAPTION | WS_THICKFRAME;
        HWND window = CreateWindowEx(0, name, name,
            flags & ~WS_VISIBLE, CW_USEDEFAULT, CW_USEDEFAULT,
            200, 200, 0, 0, GetModuleHandle(0), 0);

        return window;
    }

    return NULL;
}



int __stdcall WinMain(_In_ HINSTANCE hInstance,
    _In_ HINSTANCE hPrevInstance,
    _In_ LPSTR     lpCmdLine,
    _In_ int       nCmdShow)
{
    LOG.enable(SETTINGS.get(WallSettings::EnableDebugLog));
    LOG << L"Begin";

    // Allow only single instance of the application
    HWND oldWindow = FindWindow(APP_NAME, APP_NAME);
    if (oldWindow != NULL) {
        LOG << L"Activate previous instance and exit";
        PostMessage(oldWindow, WM_COMMAND, MENU_ID_SET_WALLPAPER, 0);
        return 0;
    }

    // Read list of images and apply wallpapers
    readWallpapers();
    setWallpapers(false);

    // Create main window for message handling and tray icon
    HWND window = createWindow(APP_NAME, WndProc);

    // Create tray icon - the only way to interact with the program
    createNotificationIcon(window, L"Double-click to set desktop wallpaper");

    // Watch for changes in the directory with images
    manageFolderWatcher(window);

    // If configured schedule periodic wallpaper updates
    if (SETTINGS.get(WallSettings::AutoChangeImage)) {
        SetTimer(window, EVENT_SET_WALLPAPER_SCHEDULED, ONE_MINUTE_MILLIS * (int)SETTINGS.get(WallSettings::AutoChangeInterval), NULL);
    }

    // Main program loop
    MSG msg;
    while (GetMessage(&msg, 0, 0, 0)) DispatchMessage(&msg);

    // Terminate folder watcher thread
    manageFolderWatcher(window, false);

    // Remove icon from tray
    deleteNotificationIcon(window);

    return 0;
}

#pragma endregion