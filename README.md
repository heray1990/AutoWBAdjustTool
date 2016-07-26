# AutoWBAdjustTool

AutoWBAdjustTool is a tool for auto white balance adjustment of some TVs. For now, it supports Letv, Haier and CAN. 

## Requirements

### Software
* OS
	* Windows (XP, Vista or 7).
* Development Tool
	* Visual Basic 6.0
	* Visual C++ 6.0
* Driver
	* FTDI USB Serial Converter Drivers (Google it to download this driver. I am using CDM 2.02.04.exe).
	* CA-SDK for CA-310/CA-210/100Plus (You can find the SDK in CD/DVD in KONICA MINOLTA products or its offical website).
	* Drivers for Chroma VPG products (To install **VPGMaster** provided by Chroma so that you can get drivers for Chroma VPG products).
	* I2CBridge.0.1.4.exe (Install it so that you can communicate with Hx6310 by I2C).

### Hardware
* A PC with softwares introduced above. 
* A TV which needs to adjust white balance.
* CA-310 or CA-210 with a USB B Type cable.
* Chroma VPG products (such as, 22294, 22294-A, 2401, 2402) with a USB B Type cable and a signal cable (for example, a HDMI cable).
* A debug tool which connects PC and TV.
* A network cable (Some TVs may use a network cable instead of a debug tool to connect to PC).

## Building