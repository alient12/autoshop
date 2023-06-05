# Autoshop
This script will automate some works in photoshop:
- import images
- save output
- make blend mode option easier
- make opacity setting easier
- output mask

This script has a simple GUI that everyone with no experience in photoshop can do data augmentation.

## requirements
- A windows os with installed photoshop (it has to be installed, portable version has not been tested yet)
- python 3 (3.8.3 is recommended, other versions have not been tested yet)

## Installation

`pywin32` required for this script

```bash
pip install pywin32
```

## Usage

change `BASE_PATH` in line 6 to absolute path that you want use.
3 folders are required in `BASE_PATH`:
- backgrounds
- foregrounds
- outputs
- masks

Put backgrounds in backgrounds folder and foregrounds in foregrounds folder. Run the script using this command.
```bash
python autoshop.py
```
Now enjoy augmentation procces :)

