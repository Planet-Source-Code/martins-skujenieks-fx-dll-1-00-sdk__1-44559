@ECHO OFF
CLS
ECHO ---------------------------------------------------------------------------
ECHO FX.DLL 1.00 SDK Setup
ECHO ---------------------------------------------------------------------------
ECHO.
ECHO FX.DLL Version 1.00
ECHO Copyright (C) Martins Skujenieks 2003
ECHO.
ECHO Visit http://www.exe.times.lv
ECHO E-Mail exe.times@e-apollo.lv
ECHO.
ECHO.
ECHO ---------------------------------------------------------------------------
ECHO Press CTRL+C to exit . . .
PAUSE
CLS
ECHO ---------------------------------------------------------------------------
ECHO End User Licence Agreement (EULA)
ECHO ---------------------------------------------------------------------------
ECHO.
ECHO This product is provided "as is", with no guarantee of completeness or
ECHO accuracy and without warranty of any kind, express or implied.
ECHO.
ECHO In no event will developer be liable for damages of any kind that may be
ECHO incurred with your hardware, peripherals or software programs.
ECHO.
ECHO You may create one copy of the product on any single computer for your
ECHO personal, non-commercial, home use only, provided you keep intact all
ECHO copyright and other proprietary notices.
ECHO.
ECHO This product and all of its parts may not be copied, emulated, cloned,
ECHO rented, leased, sold, reproduced, modified, decompiled, disassembled,
ECHO otherwise reverse engineered, republished, uploaded, posted, transmitted
ECHO or distributed in any way, without prior written consent of the developer.
ECHO.
ECHO.
ECHO ---------------------------------------------------------------------------
ECHO Press CTRL+C to exit . . .
PAUSE
CLS
ECHO ---------------------------------------------------------------------------
ECHO Installing...
ECHO ---------------------------------------------------------------------------
ECHO.
COPY fx.pkg release\fx.dll
COPY fx.pkg samples\fxdllp~1\fx.dll
COPY fx.pkg samples\window~1\fx.dll
COPY fx.pkg samples\hellow~1\fx.dll
COPY fxsdk.pkg sdk\fxsdk.exe
ECHO.
ECHO.
ECHO Installation was successfully completed...