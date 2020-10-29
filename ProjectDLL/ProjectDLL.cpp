/* 
	Created by Rizky Khapidsyah
*/

#include "stdafx.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

//Prototipe
int _stdcall Perkalian(int x, int y);

//Definisi
int _stdcall Perkalian(int x, int y){
    int Hasil;

    Hasil = x * y;

    return Hasil;
}