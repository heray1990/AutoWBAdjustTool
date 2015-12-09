
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the COLORT_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// COLORT_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef COLORT_EXPORTS
#define COLORT_API __declspec(dllexport)
#else
#define COLORT_API __declspec(dllimport)
#endif

// This class is exported from the ColorT.dll
class COLORT_API CColorT {
public:
	CColorT(void);
	// TODO: add your methods here.
};

extern COLORT_API int nColorT;

COLORT_API int fnColorT(void);

///////////////////////////////////////////////////////////////////////////////////////////////

typedef struct COLORSPEC
{
  unsigned int sx;
  unsigned int sy;
  unsigned int LimLV;
  unsigned int PriRR;
  unsigned int PriGG;
  unsigned int PriBB;
  unsigned int xt;
  unsigned int yt;
  unsigned int cxt;
  unsigned int cyt;
  unsigned int LowRR;
  unsigned int LowGG;
  unsigned int LowBB;
}COLORSPEC, *pCOLORSPEC;

typedef struct REALCOLOR
{
  unsigned int sx;
  unsigned int sy;
  unsigned int Lv;
}REALCOLOR, *pREALCOLOR;

typedef struct REALRGB
{
  unsigned int cRR;
  unsigned int cGG;
  unsigned int cBB;
}REALRGB, *pREALRGB;

  unsigned int maxColorRGB_OFF;
  unsigned int minColorRGB_OFF;
  unsigned int maxColorRGB_GAN;
  unsigned int minColorRGB_GAN;

  COLORSPEC Spec14000K;
  COLORSPEC Spec13000K;
  COLORSPEC Spec11000K;
  COLORSPEC Spec9300K;
  COLORSPEC Spec8500K;
  COLORSPEC Spec7500K;
  COLORSPEC Spec6500K;
  COLORSPEC Spec5800K;
  COLORSPEC Spec5000K;
  COLORSPEC Spec4000K;
  COLORSPEC Spec3000K;
  COLORSPEC Spec2600K;
  COLORSPEC PrimaryData;
  REALCOLOR CurrentData;
  REALRGB CalcRGB;
//  CString strBuff;

    char buf[255];
	const int nDefault=0;
	short ReadDataBuffer[255];
	unsigned int ca_x=0;
	unsigned int ca_y=0;
	unsigned int ca_lv=0;
	int min_rgb=0;
	int AdjustGAN=0;


COLORT_API int _stdcall  initColorTemp(int *pTimming, int *pPattern,int *pMaxLV, int *pMinLV, BOOL *pCalibraEN, BOOL *pMiniBriEN, char* ModelFile);
COLORT_API int _stdcall  DeinitColorTemp(char* ModelFile);
COLORT_API int _stdcall  setColorTemp(int colorTemp, pCOLORSPEC pSpecData,int GANref);
COLORT_API int _stdcall  checkColorTemp(pREALCOLOR PGetColor,int colorTemp);
COLORT_API int _stdcall  adjustColorTemp(int FixValue, BOOL xyAdjMode, BOOL AdjStep, pREALRGB pAdjRGB);

 void  delay(unsigned milliseconds);
 int   savedata(pCOLORSPEC pColorST,char* CT);
 int   getdata(pCOLORSPEC pColorST,char* CT);
 BOOL  CheckRGBisInRangeOkorNO(COLORSPEC rgb);
 void  VerifyRGB(unsigned int& RGB);

 void  AverageData(pCOLORSPEC pColorST);
 void  ReLoadRGB(int colorTemp);

//COLORT_API