// I was only able to make this thanks to https://stackoverflow.com/users/407438/jim-brissom
// and his answer here: https://stackoverflow.com/questions/3859276/get-console-handle
// as well as consulting the C# source code for ColorTool (https://github.com/Microsoft/console/tree/master/tools/ColorTool)
// and of course Microsoft Documentation - thanks everyone!

#include <sstream>
#include <iostream>
#include <fstream> // only for logfile (debugging) - could be taken out
#include <windows.h>

int main (int argc, char *argv[]) {
    // Command line arguments:
    // 0 : name of executable
    // 1 : process ID of console
    // 2-17 : colors to apply
    if (argc < 3) {
        printf("Not enough arguments!");
        return -5;
    }

    int returnCode = EXIT_SUCCESS;
    DWORD process_id = atoi(argv[1]);
    COLORREF newColorTable[16] = {0};

    int count = 0;

    while (count < argc - 2) {
        newColorTable[count++] = strtoul(argv[count + 2], NULL, 0);
    }

    /*
        0  = EXIT_SUCCESS
        1  = Could not write to log file, but may still succeed
        -1 = Failed to attach to console
        -2 = Failed to get handle of console
        -3 = Failed to get ScreenBufferInfoEx of console
        -4 = Failed to set ScreenBufferInfoEx of console
        -5 = Too few arguments supplied (minimum: [executable] PID COLOR0)
    */

    std::ofstream fs("c:\\AMD\\cpp_attacher.txt"); 
    if (!fs) {
        std::cerr<<"Cannot open the output file."<<std::endl;
        returnCode = 1;
    }
    fs<< "PROCESSID: " << process_id << "\n";
    printf("OK!\n");

    if (AttachConsole(process_id)) {
        HANDLE hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
        if (hStdOut != NULL) {
            fs<< "ConHandle: " << hStdOut << "\n";
            CONSOLE_SCREEN_BUFFER_INFOEX console_buffer_info = { 96 };

            if (GetConsoleScreenBufferInfoEx(hStdOut, &console_buffer_info)) {
                for (int i = 0; i < count; i++) {
                    console_buffer_info.ColorTable[i] = newColorTable[i];
                }
                if (! SetConsoleScreenBufferInfoEx(hStdOut, &console_buffer_info)) {
                    fs<<"Failed to set ScreenBufferInfoEx of consol\n";
                    returnCode = -4;
                }
            } else {
                fs<<"Failed to get ScreenBufferInfoEx of console\n";
                returnCode = -3;
            }
        } else {
            fs<<"Failed to get handle of console\n";
            returnCode = -2;
        }   
        FreeConsole();
    } else {
        fs<<"Failed to attach to console\n";
        returnCode = -1;
    }

    fs<<"ERRORCODE: " << returnCode;
    if (returnCode != 0) {
        fs<< "\nWIN32_LASTERROR: " << GetLastError();
    }
    fs.close();
    return returnCode;
}
