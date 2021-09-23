# ExtractInterfaceBulk
Extract Interface Bulk from Visual Studio for one whole project

# Configuration

## BaseInterface
### normal class with base interface inherited
If original class already inheirit one base interface, like "public class BaseClass ： IBaseInterface", put name "IBaseInterface" into ExtractInterfaceBulk.Execute below line "var baseInterface = "
### class without base interface inherited
Modify "Setting/IgnoreClasses.txt", each sepecial class per line

## Ignore Folders
Change above line of "BaseInterface"

## Ignore Classes
Modify "Setting/IgnoreClasses.txt", each ignored class per line

## Delete Lines from new extracted interface

### normal full match line
Modify "Setting/DeleteLines.txt", each to be deleted line per line
### lines with startwith
Modify "Setting/DeleteLinesWildCard.txt", each to be deleted line wildcard per line

## Move interfaces to be another folder
Modify variable "MoveInterfaceToFolder" inner ExtractInterfaceBulk.cs, and empty means no such action.

# How To Use
After you modify above changes and run this project, right click any project, you will see a menuitem "Extract Interface Bulk" at the bottom.

It will finish below tasks:
1, extract interface for all classes files under selected Project
2, modify interface files while processed class has inheirited base interface
3, remove lines with full match to startwith from interface files
4, move interface file to one specific folder instead of same as class file if you want.