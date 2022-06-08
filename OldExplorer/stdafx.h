#pragma once
#include "targetver.h"

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used parts from the header
#define OEMRESOURCE
//////////////////////////////////////////////////////////////////////
#include "windowsx.h"
#include "Resource.h"
#include <atlbase.h>			//A2T T2A W2T T2W A2W W2A
#include <sstream>
#include <variant>
#include <fstream>
#include <exdispid.h>
#include <mshtmdid.h>
#include <set>
#include <future>				// std::async, std::future
#include <filesystem>			//filesystem
#include <Commdlg.h>
#include "time.h"
#include <bitset>
#include <DispatcherQueue.h>
#include <regex>
#include <ShObjIdl_core.h>
#include <pathcch.h>
#include <Shellapi.h>
////////////////////////////////////////////////////////////////////// Sound control
#include <mmeapi.h>
#pragma comment (lib, "winmm.lib")
////////////////////////////////////////////////////////////////////// Magnification
#include "magnification.h"
#pragma comment(lib, "Magnification.lib")
////////////////////////////////////////////////////////////////////// WindowsTools
#include "uxtheme.h"
#pragma comment (lib, "UxTheme.lib")
#include "WindowsTools.h"
