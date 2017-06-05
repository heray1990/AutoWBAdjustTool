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

COLORT_API int _stdcall ColorTInit(pudtConfigData pConfigdata)
{
	ConfigData.intCLEVELRGBOMax=pConfigdata->intCLEVELRGBOMax;
	ConfigData.intCLEVELRGBOMin=pConfigdata->intCLEVELRGBOMin;	
	ConfigData.intCLEVELRGBGMax=pConfigdata->intCLEVELRGBGMax;
	ConfigData.intCLEVELRGBGMin=pConfigdata->intCLEVELRGBGMin;
	ConfigData.intSPECCool1x=pConfigdata->intSPECCool1x;
	ConfigData.intSPECCool1y=pConfigdata->intSPECCool1y;
	ConfigData.intSPECCool1Lv=pConfigdata->intSPECCool1Lv;
	ConfigData.intPRESETGANCool1R=pConfigdata->intPRESETGANCool1R;
	ConfigData.intPRESETGANCool1G=pConfigdata->intPRESETGANCool1G;
	ConfigData.intPRESETGANCool1B=pConfigdata->intPRESETGANCool1B;
	ConfigData.intPRESETOFFCool1R=pConfigdata->intPRESETOFFCool1R;
	ConfigData.intPRESETOFFCool1G=pConfigdata->intPRESETOFFCool1G;
	ConfigData.intPRESETOFFCool1B=pConfigdata->intPRESETOFFCool1B;
	ConfigData.intTOLCool1xt=pConfigdata->intTOLCool1xt;
	ConfigData.intTOLCool1yt=pConfigdata->intTOLCool1yt;
	ConfigData.intCHKCool1Cxt=pConfigdata->intCHKCool1Cxt;
	ConfigData.intCHKCool1Cyt=pConfigdata->intCHKCool1Cyt;
	ConfigData.intSPECNormalx=pConfigdata->intSPECNormalx;
	ConfigData.intSPECNormaly=pConfigdata->intSPECNormaly;
	ConfigData.intSPECNormalLv=pConfigdata->intSPECNormalLv;
	ConfigData.intPRESETGANNormalR=pConfigdata->intPRESETGANNormalR;
	ConfigData.intPRESETGANNormalG=pConfigdata->intPRESETGANNormalG;
	ConfigData.intPRESETGANNormalB=pConfigdata->intPRESETGANNormalB;
	ConfigData.intPRESETOFFNormalR=pConfigdata->intPRESETOFFNormalR;
	ConfigData.intPRESETOFFNormalG=pConfigdata->intPRESETOFFNormalG;
	ConfigData.intPRESETOFFNormalB=pConfigdata->intPRESETOFFNormalB;
	ConfigData.intTOLNormalxt=pConfigdata->intTOLNormalxt;
	ConfigData.intTOLNormalyt=pConfigdata->intTOLNormalyt;
	ConfigData.intCHKNormalCxt=pConfigdata->intCHKNormalCxt;
	ConfigData.intCHKNormalCyt=pConfigdata->intCHKNormalCyt;
	ConfigData.intSPECWarm1x=pConfigdata->intSPECWarm1x;
	ConfigData.intSPECWarm1y=pConfigdata->intSPECWarm1y;
	ConfigData.intSPECWarm1Lv=pConfigdata->intSPECWarm1Lv;
	ConfigData.intPRESETGANWarm1R=pConfigdata->intPRESETGANWarm1R;
	ConfigData.intPRESETGANWarm1G=pConfigdata->intPRESETGANWarm1G;
	ConfigData.intPRESETGANWarm1B=pConfigdata->intPRESETGANWarm1B;
	ConfigData.intPRESETOFFWarm1R=pConfigdata->intPRESETOFFWarm1R;
	ConfigData.intPRESETOFFWarm1G=pConfigdata->intPRESETOFFWarm1G;
	ConfigData.intPRESETOFFWarm1B=pConfigdata->intPRESETOFFWarm1B;
	ConfigData.intTOLWarm1xt=pConfigdata->intTOLWarm1xt;
	ConfigData.intTOLWarm1yt=pConfigdata->intTOLWarm1yt;
	ConfigData.intCHKWarm1Cxt=pConfigdata->intCHKWarm1Cxt;
	ConfigData.intCHKWarm1Cyt=pConfigdata->intCHKWarm1Cyt;
	ConfigData.intMAGICVALGMin=pConfigdata->intMAGICVALGMin;
	ConfigData.intMAGICVALOMin=pConfigdata->intMAGICVALOMin;
	ConfigData.intMAGICVALGMax=pConfigdata->intMAGICVALGMax;
	ConfigData.intMAGICVALOMax=pConfigdata->intMAGICVALOMax;


	maxColorRGB_OFF = pConfigdata->intCLEVELRGBOMax;
	minColorRGB_OFF = pConfigdata->intCLEVELRGBOMin;	
	maxColorRGB_GAN = pConfigdata->intCLEVELRGBGMax;
	minColorRGB_GAN = pConfigdata->intCLEVELRGBGMin;

	Getdata(&SpecCool1,&ConfigData,"COOL1");
	Getdata(&SpecNormal,&ConfigData,"NORMAL");
	Getdata(&SpecWarm1,&ConfigData,"WARM1");	

    return true;
}

COLORT_API int _stdcall ColorTDeInit(pudtConfigData pConfigdata)
{
	Savedata(&SpecCool1,&ConfigData,"COOL1");
	Savedata(&SpecNormal,&ConfigData,"NORMAL");
	Savedata(&SpecWarm1,&ConfigData,"WARM1");
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

int Getdata(pCOLORSPEC pColorST,pudtConfigData pConfigdata,char* CT)
{
	
	if (strcmp(CT, "COOL1") == 0)
	{
		pColorST->sx =pConfigdata->intSPECCool1x;
		pColorST->sy =pConfigdata->intSPECCool1y;
		pColorST->LimLV =pConfigdata->intSPECCool1Lv;
		pColorST->PriRR =pConfigdata->intPRESETGANCool1R;
		pColorST->PriGG =pConfigdata->intPRESETGANCool1G;
		pColorST->PriBB =pConfigdata->intPRESETGANCool1B;
		pColorST->LowRR =pConfigdata->intPRESETOFFCool1R;
		pColorST->LowGG =pConfigdata->intPRESETOFFCool1G;
		pColorST->LowBB =pConfigdata->intPRESETOFFCool1B;
		pColorST->xt =pConfigdata->intTOLCool1xt;
		pColorST->yt =pConfigdata->intTOLCool1yt;
		pColorST->cxt =pConfigdata->intCHKCool1Cxt;
		pColorST->cyt =pConfigdata->intCHKCool1Cyt;
	}

	if (strcmp(CT, "NORMAL") == 0)
	{
		pColorST->sx =pConfigdata->intSPECNormalx;
		pColorST->sy =pConfigdata->intSPECNormaly;
		pColorST->LimLV =pConfigdata->intSPECNormalLv;
		pColorST->PriRR =pConfigdata->intPRESETGANNormalR;
		pColorST->PriGG =pConfigdata->intPRESETGANNormalG;
		pColorST->PriBB =pConfigdata->intPRESETGANNormalB;
		pColorST->LowRR =pConfigdata->intPRESETOFFNormalR;
		pColorST->LowGG =pConfigdata->intPRESETOFFNormalG;
		pColorST->LowBB =pConfigdata->intPRESETOFFNormalB;
		pColorST->xt =pConfigdata->intTOLNormalxt;
		pColorST->yt =pConfigdata->intTOLNormalyt;
		pColorST->cxt =pConfigdata->intCHKNormalCxt;
		pColorST->cyt =pConfigdata->intCHKNormalCyt;
	}

	if (strcmp(CT, "WARM1") == 0)
	{
		pColorST->sx =pConfigdata->intSPECWarm1x;
		pColorST->sy =pConfigdata->intSPECWarm1y;
		pColorST->LimLV =pConfigdata->intSPECWarm1Lv;
		pColorST->PriRR =pConfigdata->intPRESETGANWarm1R;
		pColorST->PriGG =pConfigdata->intPRESETGANWarm1G;
		pColorST->PriBB =pConfigdata->intPRESETGANWarm1B;
		pColorST->LowRR =pConfigdata->intPRESETOFFWarm1R;
		pColorST->LowGG =pConfigdata->intPRESETOFFWarm1G;
		pColorST->LowBB =pConfigdata->intPRESETOFFWarm1B;
		pColorST->xt =pConfigdata->intTOLWarm1xt;
		pColorST->yt =pConfigdata->intTOLWarm1yt;
		pColorST->cxt =pConfigdata->intCHKWarm1Cxt;
		pColorST->cyt =pConfigdata->intCHKWarm1Cyt;
	}
	pColorST->MagicValXStepGain = pConfigdata->intMAGICVALGMin;
	pColorST->MagicValXStepOffset = pConfigdata->intMAGICVALOMin;
	pColorST->MagicValYStepGain = pConfigdata->intMAGICVALGMax;
	pColorST->MagicValYStepOffset = pConfigdata->intMAGICVALOMax;
	
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



int Savedata(pCOLORSPEC pColorST,pudtConfigData pConfigdata,char* CT)
{
	if(strcmp(CT, "COOL1") == 0)
	{
		pConfigdata->intPRESETGANCool1R=pColorST->PriRR;
		pConfigdata->intPRESETGANCool1G=pColorST->PriGG;
		pConfigdata->intPRESETGANCool1B=pColorST->PriBB;
		pConfigdata->intPRESETOFFCool1R=pColorST->LowRR;
		pConfigdata->intPRESETOFFCool1G=pColorST->LowGG;
		pConfigdata->intPRESETOFFCool1B=pColorST->LowBB;
	}
	if (strcmp(CT, "NORMAL") == 0)
	{
		pConfigdata->intPRESETGANNormalR=pColorST->PriRR;
		pConfigdata->intPRESETGANNormalG=pColorST->PriGG;
		pConfigdata->intPRESETGANNormalB=pColorST->PriBB;
		pConfigdata->intPRESETOFFNormalR=pColorST->LowRR;
		pConfigdata->intPRESETOFFNormalG=pColorST->LowGG;
		pConfigdata->intPRESETOFFNormalB=pColorST->LowBB;	
	}
	if (strcmp(CT, "WARM1") == 0)
	{
		pConfigdata->intPRESETGANWarm1R=pColorST->PriRR;
		pConfigdata->intPRESETGANWarm1G=pColorST->PriGG;
		pConfigdata->intPRESETGANWarm1B=pColorST->PriBB;
		pConfigdata->intPRESETOFFWarm1R=pColorST->LowRR;
		pConfigdata->intPRESETOFFWarm1G=pColorST->LowGG;
		pConfigdata->intPRESETOFFWarm1B=pColorST->LowBB;	
	}
	return true;
}