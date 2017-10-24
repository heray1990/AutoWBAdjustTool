// ColorT.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "ColorT.h"
#include "stdlib.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

// This is the constructor of a class that has been exported.
// see ColorT.h for the class definition
CColorT::CColorT()
{ 
	return; 
}

COLORT_API int _stdcall ColorTInit(char* ModelFile,char* pCurDir)
{
	//::GetCurrentDirectory(512,buf);
	strcpy(buf, pCurDir);
	strcat(buf, "\\");
	strcat(buf, ModelFile);
	strcat(buf, "\\CONFIG.ini");
  
	maxColorRGB_OFF = GetPrivateProfileInt("Color_Level_RGB_OFF", "####max", nDefault, buf);
	minColorRGB_OFF = GetPrivateProfileInt("Color_Level_RGB_OFF", "####min", nDefault, buf);	
	maxColorRGB_GAN = GetPrivateProfileInt("Color_Level_RGB_GAN", "####max", nDefault, buf);
	minColorRGB_GAN = GetPrivateProfileInt("Color_Level_RGB_GAN", "####min", nDefault, buf);
	

	// COOL1
	getdata(&SpecCool1, "COOL1");
	// NORMAL
    getdata(&SpecNormal, "NORMAL");
	// WARM1
    getdata(&SpecWarm1, "WARM1");

    return true;
}

COLORT_API int _stdcall ColorTDeInit()
{
	savedata(&SpecCool1, "COOL1");
	savedata(&SpecNormal, "NORMAL");
	savedata(&SpecWarm1, "WARM1");

    return true;
}

COLORT_API int _stdcall ColorTSetSpec(char* colorTemp, pCOLORSPEC pSpecData,int GANref)
{
	if (strcmp(colorTemp, "COOL1") == 0)
	{
		PrimaryData = SpecCool1;
	}
	else if (strcmp(colorTemp, "NORMAL") == 0)
	{
		PrimaryData = SpecNormal;
	}
	else if (strcmp(colorTemp, "WARM1") == 0)
	{
		PrimaryData = SpecWarm1; 
	}

	AdjustGAN = GANref;    
	pSpecData->sx = PrimaryData.sx;
	pSpecData->sy = PrimaryData.sy;
	pSpecData->LimLV = PrimaryData.LimLV;

	if (AdjustGAN == 1)
	{
		pSpecData->PriRR = CalcRGB.cRR = PrimaryData.PriRR;
		pSpecData->PriGG = CalcRGB.cGG = PrimaryData.PriGG;
		pSpecData->PriBB = CalcRGB.cBB = PrimaryData.PriBB;
	}
	else
	{
        pSpecData->PriRR = CalcRGB.cRR = PrimaryData.LowRR;
	    pSpecData->PriGG = CalcRGB.cGG = PrimaryData.LowGG;
        pSpecData->PriBB = CalcRGB.cBB = PrimaryData.LowBB;
	}
	pSpecData->xt = PrimaryData.xt;
    pSpecData->yt = PrimaryData.yt;
//	CalcGainRx=GainRx;
//	CalcGainRy=GainRy;
//	CalcGainGx=GainGx;
//	CalcGainGy=GainGy;
//	CalcGainBx=GainBx;
//	CalcGainBy=GainBy;
//    delay(40);

//    delay(tolTime);

	if (TRUE == CheckRGBisInRangeOkorNO(PrimaryData))
	{
        return true;
    }
    return CalcRGB.cRR;
}

COLORT_API int _stdcall ColorTChk(pREALCOLOR pGetColor,char* colorTemp)
{
	ca_x = pGetColor->sx;
	ca_y = pGetColor->sy;
	ca_lv = pGetColor->Lv;

	PrimaryData.PriRR = CalcRGB.cRR;           //For stepbystep adjust.
	PrimaryData.PriGG = CalcRGB.cGG;
	PrimaryData.PriBB = CalcRGB.cBB;

	if ((ca_x < PrimaryData.sx - PrimaryData.cxt) ||
		(ca_x > PrimaryData.sx + PrimaryData.cxt))
	{
		if ((ca_y < PrimaryData.sy - PrimaryData.cyt) ||
			(ca_y > PrimaryData.sy + PrimaryData.cyt))
		{
			return 0;
		}
		else
		{
			return 1;
		}
	}
	else
	{
		if ((ca_y < PrimaryData.sy - PrimaryData.cyt) ||
			(ca_y > PrimaryData.sy + PrimaryData.cyt))
		{
			return 2;
		}
	}

	ReLoadRGB(colorTemp);
	return 3;
}

void AverageData(pCOLORSPEC pColorST)
{
	if (AdjustGAN == 1)
	{
        pColorST->PriRR = (pColorST->PriRR + CalcRGB.cRR) / 2;
        //pColorST->PriGG = (pColorST->PriGG + CalcRGB.cGG) / 2;
		//pColorST->PriGG = 128;
        pColorST->PriBB = (pColorST->PriBB + CalcRGB.cBB) / 2;
	}
	else
	{
        pColorST->LowRR = (pColorST->LowRR + CalcRGB.cRR) / 2;
        pColorST->LowGG = (pColorST->LowGG + CalcRGB.cGG) / 2;
        pColorST->LowBB = (pColorST->LowBB + CalcRGB.cBB) / 2;
	}
}

void ReLoadRGB(char* colorTemp)
{
	if (strcmp(colorTemp, "COOL1") == 0)
	{
		AverageData(&SpecCool1);
	}
	else if (strcmp(colorTemp, "NORMAL") == 0)
	{
		AverageData(&SpecNormal);
	}
	else if (strcmp(colorTemp, "WARM1") == 0)
	{
		AverageData(&SpecWarm1); 
	}
}

COLORT_API int _stdcall  ColorTAdjRGBGain(pREALRGB pAdjRGB)
{
	if (ca_y < PrimaryData.sy - PrimaryData.yt)
	{
		CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sy - ca_y) / PrimaryData.MagicValYStepGain;
	}
	else
	{
		if (ca_y > PrimaryData.sy + PrimaryData.yt)
		{
			CalcRGB.cBB = PrimaryData.PriBB + (ca_y - PrimaryData.sy) / PrimaryData.MagicValYStepGain;
		}
		else
		{
			if (ca_x > PrimaryData.sx + PrimaryData.xt)
			{
				CalcRGB.cRR = PrimaryData.PriRR - (ca_x - PrimaryData.sx) / PrimaryData.MagicValXStepGain;
			}
			else
			{
				if (ca_x < PrimaryData.sx - PrimaryData.xt)
				{
					CalcRGB.cRR = PrimaryData.PriRR + (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;
				}
			}
		}
	}

    VerifyRGB(CalcRGB.cRR);
    VerifyRGB(CalcRGB.cBB);
	pAdjRGB->cRR = CalcRGB.cRR;
    pAdjRGB->cGG = CalcRGB.cGG;
	pAdjRGB->cBB = CalcRGB.cBB;

    return true;
}

/////////////////////////////////////////////////////////////////////////////
////////////////////////Add for Colortemperature App.////////////////////////
/////////////////////////////////////////////////////////////////////////////
COLORT_API int _stdcall  ColorTAdjRGBGainLetv(int FixValue, pREALRGB pAdjRGB, int *pResultCode)
{
	switch (FixValue)
	{
		case 1:
			if (ca_y < PrimaryData.sy - PrimaryData.yt)
			{
				CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sy - ca_y) / PrimaryData.MagicValYStepGain;
				*pResultCode = 3;
				break;
			}
			if (ca_x < PrimaryData.sx - PrimaryData.xt)
			{
				CalcRGB.cRR = PrimaryData.PriRR + (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;

				if (CalcRGB.cRR >= 135)
				{
					/* No matter x is OK or not, go to case 2, 
					   so be careful about x when adjusting G 
					   gain in case 2. */
					CalcRGB.cRR = 135;
					*pResultCode = 2;
				}
			}
			else if (ca_x > PrimaryData.sx + PrimaryData.xt)
			{
				CalcRGB.cRR = PrimaryData.PriRR - (ca_x - PrimaryData.sx) / PrimaryData.MagicValXStepGain;
			}
			else    // x is OK
			{
				*pResultCode = 2;
			}
			break;
		case 2:
			if (CalcRGB.cGG > 128)
			{
				CalcRGB.cGG = 128;
			}
			// Adjust G Gain to match y.
			if (ca_y < PrimaryData.sy - PrimaryData.yt)
			{
				CalcRGB.cGG = CalcRGB.cGG + (PrimaryData.sy - ca_y) / PrimaryData.MagicValYStepGain;
			}
			else if (ca_y > PrimaryData.sy + PrimaryData.yt)
			{
				CalcRGB.cGG = CalcRGB.cGG - (ca_y - PrimaryData.sy) / PrimaryData.MagicValYStepGain;
			}

			if (ca_x > PrimaryData.sx + PrimaryData.xt)
			{
				CalcRGB.cRR = CalcRGB.cRR - (ca_x - PrimaryData.sx) / PrimaryData.MagicValXStepGain;
			}
			else if (ca_x < PrimaryData.sx - PrimaryData.xt)
			{
				CalcRGB.cRR = CalcRGB.cRR + (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;

				if (CalcRGB.cRR >= 135)
				{
					/* R gain is saturated, but x is too small. 
					   Go to case 4 to fix it (decrease B gain). */
					CalcRGB.cRR = 135;
					*pResultCode = 4;

					CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;
				}
			}
			break;
		case 3:   //  "normal adjust" 
			if (ca_y < PrimaryData.sy - PrimaryData.yt)
			{
				CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sy - ca_y) / PrimaryData.MagicValYStepGain;
			}
			else
			{
				if (ca_y > PrimaryData.sy + PrimaryData.yt)
				{
					CalcRGB.cBB = PrimaryData.PriBB + (ca_y - PrimaryData.sy) / PrimaryData.MagicValYStepGain;
				
					if (CalcRGB.cBB >= 135)
					{
						CalcRGB.cBB = 135;
						*pResultCode = 1;
					}
				}
				else
				{
					if (ca_x > PrimaryData.sx + PrimaryData.xt)
					{
						CalcRGB.cRR = PrimaryData.PriRR - (ca_x - PrimaryData.sx) / PrimaryData.MagicValXStepGain;
					}
					else
					{
						if (ca_x < PrimaryData.sx - PrimaryData.xt)
						{
							CalcRGB.cRR = PrimaryData.PriRR + (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;
						
							if (CalcRGB.cRR >= 135)
							{
								/* R gain is saturated, but x is too small. 
								   Go to case 4 to fix it (decrease B gain). */
								CalcRGB.cRR = 135;
								*pResultCode = 4;

								CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;
							}
						}
					}
				}
			}
			break;
		case 4:
			if (ca_x < PrimaryData.sx - PrimaryData.xt)
			{
				CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepGain;
			}
			else if (ca_x > PrimaryData.sx + PrimaryData.xt)
			{
				CalcRGB.cBB = PrimaryData.PriBB + (ca_x - PrimaryData.sx) / PrimaryData.MagicValXStepGain;
			}
			else
			{
				if (ca_y < PrimaryData.sy - PrimaryData.yt)
				{
					CalcRGB.cGG = CalcRGB.cGG + (PrimaryData.sy - ca_y) / PrimaryData.MagicValYStepGain;

					if (CalcRGB.cGG > 128)
					{
						CalcRGB.cGG = 128;
					}
				}
				else if (ca_y > PrimaryData.sy + PrimaryData.yt)
				{
					CalcRGB.cGG = CalcRGB.cGG - (ca_y - PrimaryData.sy) / PrimaryData.MagicValYStepGain;
				}

				*pResultCode = 2;
			}
			break;
		default:
			break;
	}

    VerifyRGB(CalcRGB.cRR);
    VerifyRGB(CalcRGB.cBB);
	pAdjRGB->cRR = CalcRGB.cRR;
    pAdjRGB->cGG = CalcRGB.cGG;
	pAdjRGB->cBB = CalcRGB.cBB;

    return true;
}

COLORT_API int _stdcall  ColorTAdjRGBOffset(pREALRGB pAdjRGB)
{
	if (ca_y < PrimaryData.sy - PrimaryData.yt)
	{
		CalcRGB.cBB = PrimaryData.PriBB - (PrimaryData.sy - ca_y) / PrimaryData.MagicValYStepOffset;
	}
	else
	{
		if (ca_y > PrimaryData.sy + PrimaryData.yt)
		{
			CalcRGB.cBB = PrimaryData.PriBB + (ca_y - PrimaryData.sy) / PrimaryData.MagicValYStepOffset;
		}
		else
		{
			if (ca_x > PrimaryData.sx + PrimaryData.xt)
			{
				CalcRGB.cRR = PrimaryData.PriRR - (ca_x - PrimaryData.sx) / PrimaryData.MagicValXStepOffset;
			}
			else
			{
				if (ca_x < PrimaryData.sx - PrimaryData.xt)
				{
					CalcRGB.cRR = PrimaryData.PriRR + (PrimaryData.sx - ca_x) / PrimaryData.MagicValXStepOffset;
				}
			}
		}
	}

    VerifyRGB(CalcRGB.cRR);
    VerifyRGB(CalcRGB.cBB);
	pAdjRGB->cRR = CalcRGB.cRR;
    pAdjRGB->cGG = CalcRGB.cGG;
	pAdjRGB->cBB = CalcRGB.cBB;

    return true;
}

void delay(unsigned milliseconds)
{
	Sleep(milliseconds);
}

int savedata(pCOLORSPEC pColorST,char* CT)
{
	char strTemp[18];
	char preset[32] = "PRESET_GAN_";
	char lowset[32] = "PRESET_OFF_";

    strcat(preset, CT);
    strcat(lowset, CT);

    if (0 != pColorST->PriRR)
	{
		if (AdjustGAN == 1)
		{
	        WritePrivateProfileString(preset, "###R", _itoa(pColorST->PriRR, strTemp, 10), buf);
     	    WritePrivateProfileString(preset, "###G", _itoa(pColorST->PriGG, strTemp, 10), buf);
	        WritePrivateProfileString(preset, "###B", _itoa(pColorST->PriBB, strTemp, 10), buf);	
		}
		else
		{
	        WritePrivateProfileString(lowset, "###R", _itoa(pColorST->LowRR, strTemp, 10), buf);
     	    WritePrivateProfileString(lowset, "###G", _itoa(pColorST->LowGG, strTemp, 10), buf);
	        WritePrivateProfileString(lowset, "###B", _itoa(pColorST->LowBB, strTemp, 10), buf);	
		}
	}
	return true;
}


BOOL CheckRGBisInRangeOkorNO(COLORSPEC rgb)
{
	if (AdjustGAN == 1)
	{
		if (rgb.PriRR < minColorRGB_GAN
			|| rgb.PriRR > maxColorRGB_GAN
			|| rgb.PriGG < minColorRGB_GAN
			|| rgb.PriGG > maxColorRGB_GAN
			|| rgb.PriBB < minColorRGB_GAN
			|| rgb.PriBB > maxColorRGB_GAN)
			return false;
	    else
			return TRUE;
	}
	else
	{
	    if (rgb.LowRR < minColorRGB_OFF
			|| rgb.LowRR > maxColorRGB_OFF
			|| rgb.LowGG < minColorRGB_OFF
			|| rgb.LowGG > maxColorRGB_OFF
			|| rgb.LowBB < minColorRGB_OFF
			|| rgb.LowBB > maxColorRGB_OFF)
			return false;
	    else
			return TRUE;
	}
}

void VerifyRGB(int& RGB)
{
	if (AdjustGAN == 1)
	{
	    if (RGB < minColorRGB_GAN)
    		RGB = minColorRGB_GAN;
    	else
			if (RGB > maxColorRGB_GAN) RGB = maxColorRGB_GAN;
	}
	else
	{
	    if (RGB < minColorRGB_OFF)
			RGB = minColorRGB_OFF;
	    else
			if (RGB > maxColorRGB_OFF) RGB = maxColorRGB_OFF;
	}
}

int getdata(pCOLORSPEC pColorST, char* CT)
{
//	char strTemp[16];
	char spec[32] = "SPEC_";
	char preset[32] = "PRESET_GAN_";
    char tol[32] = "TOL_";
	char chk[32] = "CHK_";
	char lowset[32] = "PRESET_OFF_";
	char magicValX[32] = "MAGIC_VAL_X";
	char magicValY[32] = "MAGIC_VAL_Y";

    strcat(spec, CT);
    strcat(preset, CT);
	strcat(tol, CT);
	strcat(chk, CT);
    strcat(lowset, CT);
    pColorST->sx = GetPrivateProfileInt(spec, "##x", nDefault,buf);
    pColorST->sy = GetPrivateProfileInt(spec, "##y", nDefault, buf);
    pColorST->LimLV = GetPrivateProfileInt(spec, "##Lv", nDefault, buf);
	pColorST->PriRR = GetPrivateProfileInt(preset, "###R", nDefault, buf);
    pColorST->PriGG = GetPrivateProfileInt(preset, "###G", nDefault, buf);
    pColorST->PriBB = GetPrivateProfileInt(preset, "###B", nDefault, buf);
	pColorST->xt = GetPrivateProfileInt(tol, "###x", nDefault, buf);
    pColorST->yt = GetPrivateProfileInt(tol, "###y", nDefault, buf);
	pColorST->cxt = GetPrivateProfileInt(chk, "###x", nDefault, buf);
    pColorST->cyt = GetPrivateProfileInt(chk, "###y", nDefault, buf);
	pColorST->LowRR = GetPrivateProfileInt(lowset, "###R", nDefault, buf);
    pColorST->LowGG = GetPrivateProfileInt(lowset, "###G", nDefault, buf);
    pColorST->LowBB = GetPrivateProfileInt(lowset, "###B", nDefault, buf);
	pColorST->MagicValXStepGain = GetPrivateProfileInt(magicValX, "#####STEP_GAIN", nDefault, buf);
	pColorST->MagicValXStepOffset = GetPrivateProfileInt(magicValX, "#####STEP_OFFSET", nDefault, buf);
	pColorST->MagicValYStepGain = GetPrivateProfileInt(magicValY, "#####STEP_GAIN", nDefault, buf);
	pColorST->MagicValYStepOffset = GetPrivateProfileInt(magicValY, "#####STEP_OFFSET", nDefault, buf);

	if (pColorST->MagicValXStepGain > pColorST->xt)
	{
		pColorST->MagicValXStepGain = pColorST->xt;
	}

	if (pColorST->MagicValXStepOffset > pColorST->xt)
	{
		pColorST->MagicValXStepOffset = pColorST->xt;
	}

	if (pColorST->MagicValYStepGain > pColorST->yt)
	{
		pColorST->MagicValYStepGain = pColorST->yt;
	}

	if (pColorST->MagicValYStepOffset > pColorST->yt)
	{
		pColorST->MagicValYStepOffset = pColorST->yt;
	}

	return true;
}