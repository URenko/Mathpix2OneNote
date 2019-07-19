from ctypes import *
from ctypes.wintypes import *
import base64
import json
from io import BytesIO

import requests
from PIL import ImageGrab

WNDPROCTYPE = WINFUNCTYPE(LONG, HWND, UINT, WPARAM, LPARAM)

wcslen = cdll.msvcrt.wcslen
wcslen.argtypes = [c_wchar_p]
wcslen.restype = UINT

DefWindowProc = ctypes.windll.user32.DefWindowProcW
DefWindowProc.restype = LONG
DefWindowProc.argtypes = [HWND, UINT, WPARAM, LPARAM]

OpenClipboard = windll.user32.OpenClipboard
OpenClipboard.argtypes = [HWND]
OpenClipboard.restype = BOOL

CloseClipboard = windll.user32.CloseClipboard
CloseClipboard.argtypes = []
CloseClipboard.restype = BOOL

EmptyClipboard = windll.user32.EmptyClipboard
EmptyClipboard.argtypes = []
EmptyClipboard.restype = BOOL

GetClipboardData = windll.user32.GetClipboardData
GetClipboardData.argtypes = [UINT]
GetClipboardData.restype = HANDLE

SetClipboardData = windll.user32.SetClipboardData
SetClipboardData.argtypes = [UINT, HANDLE]
SetClipboardData.restype = HANDLE

GlobalAlloc = windll.kernel32.GlobalAlloc
GlobalAlloc.argtypes = [UINT, c_size_t]
GlobalAlloc.restype = HGLOBAL

GlobalLock = windll.kernel32.GlobalLock
GlobalLock.argtypes = [HGLOBAL]
GlobalLock.restype = LPVOID

GlobalUnlock = windll.kernel32.GlobalUnlock
GlobalUnlock.argtypes = [HGLOBAL]
GlobalUnlock.restype = BOOL

WS_OVERLAPPEDWINDOW = 0xcf0000

CS_HREDRAW = 2
CS_VREDRAW = 1

CW_USEDEFAULT = 0x80000000

WM_DESTROY = 2
WM_CLIPBOARDUPDATE = 0x031D

WHITE_BRUSH = 0

GMEM_MOVEABLE = 2

class WNDCLASSEX(Structure):
    _fields_ = [("cbSize", c_uint),
                ("style", c_uint),
                ("lpfnWndProc", WNDPROCTYPE),
                ("cbClsExtra", c_int),
                ("cbWndExtra", c_int),
                ("hInstance", HANDLE),
                ("hIcon", HANDLE),
                ("hCursor", HANDLE),
                ("hBrush", HANDLE),
                ("lpszMenuName", LPCWSTR),
                ("lpszClassName", LPCWSTR),
                ("hIconSm", HANDLE)]

def mmml(mml):
    # The MathMLs Mathpix returned wrap brackets with <mo>, which is ugly in OneNote. <mfenced> is better.
    # TODO: Adapt all brackets, like [] {} <> ... But be careful about others like Dirac's bracket <| and |> .
    # https://developer.mozilla.org/en-US/docs/Web/MathML/Element/mfenced
    mml = mml.replace('<mo>(</mo>', '<mfenced>')
    mml = mml.replace('<mo>)</mo>', '</mfenced>')
    return mml[:5] + ' xmlns="http://www.w3.org/1998/Math/MathML"' + mml[5:]

def PyWndProcedure(hWnd, Msg, wParam, lParam):
    if Msg == WM_CLIPBOARDUPDATE:
        print('WM_CLIPBOARDUPDATE')
        OpenClipboard(None)
        # For debug
        # a = windll.user32.EnumClipboardFormats(0)
        # while a != 0:
            # print('Format:', a)
            # a = windll.user32.EnumClipboardFormats(a)
        hglb = GetClipboardData(49436)
        print(c_wchar_p(hglb).value)
        if 2 != windll.user32.EnumClipboardFormats(0):
            CloseClipboard()
        else:
            CloseClipboard()
            pic = ImageGrab.grabclipboard()
            if pic:
                _buffer = BytesIO()
                pic.save(_buffer, format='PNG')
                image_uri = "data:image/png;base64," + base64.b64encode(_buffer.getvalue()).decode('ascii')
                r = requests.post("https://api.mathpix.com/v3/latex",
                    data=json.dumps({'src': image_uri, 'formats': ['mathml'],
                        'metadata':{'count':123, 'platform':'windows 10', 'skip_recrop':'True', 'user_id':'sccre7rzloa4a9yzu9mbnlo5s6xl5h81', 'version':'snip.windows@01.02.0031'}}),
                    headers={"app_id": "mathpix_chrome", "app_key": "85948264c5d443573286752fbe8df361",
                        "User-Agent": "Mozilla/5.0", "Content-type": "application/json"})
                res = mmml(json.loads(r.text)['mathml'])
                print(res)
                OpenClipboard(None)
                EmptyClipboard()
                count = wcslen(res) + 1
                handle = GlobalAlloc(GMEM_MOVEABLE, count*sizeof(c_wchar))
                locked_handle = GlobalLock(handle)
                memmove(c_wchar_p(locked_handle), c_wchar_p(res), count*sizeof(c_wchar))
                GlobalUnlock(handle)
                SetClipboardData(49436, handle)
                CloseClipboard()
    elif Msg == WM_DESTROY:
        windll.user32.RemoveClipboardFormatListener(hWnd)
        windll.user32.PostQuitMessage(0)
    else:
        return DefWindowProc(hWnd, Msg, wParam, lParam)
    return 0
  


def main():
    WndProc = WNDPROCTYPE(PyWndProcedure)
    hInst = windll.kernel32.GetModuleHandleW(0)
    wclassName = 'Mathpix2OneNote'
    wname = 'Mathpix2OneNote'
    
    wndClass = WNDCLASSEX()
    wndClass.cbSize = sizeof(WNDCLASSEX)
    wndClass.style = CS_HREDRAW | CS_VREDRAW
    wndClass.lpfnWndProc = WndProc
    wndClass.cbClsExtra = 0
    wndClass.cbWndExtra = 0
    wndClass.hInstance = hInst
    wndClass.hIcon = 0
    wndClass.hCursor = 0
    wndClass.hBrush = windll.gdi32.GetStockObject(WHITE_BRUSH)
    wndClass.lpszMenuName = 0
    wndClass.lpszClassName = wclassName
    wndClass.hIconSm = 0
    
    regRes = windll.user32.RegisterClassExW(byref(wndClass))
    
    hWnd = windll.user32.CreateWindowExW(
        0, wclassName, wname,
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        300, 300, 0, 0, hInst, 0)
    
    if not hWnd:
        print('Failed to create window.')
        exit(0)
    
    if not windll.user32.AddClipboardFormatListener(hWnd):
        print('AddClipboardFormatListener failed.')
        exit(0)
    
    msg = MSG()
    lpmsg = pointer(msg)

    while windll.user32.GetMessageW(lpmsg, 0, 0, 0) != 0:
        windll.user32.TranslateMessage(lpmsg)
        windll.user32.DispatchMessageW(lpmsg)
    
    
if __name__ == "__main__":
    main()