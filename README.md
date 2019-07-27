# SatelliteChaserComm

These are the communication functions as used by Satellite Chaser.
Currently there’s a working ASCOM implementation and an untested Paramount implementation.
If you want to add support for a different communication protocol, you should

- Create a new Sub that connects the mount. This sub should set the variable mountconnected to true or false in a similar fashion as the ASCOM connection Sub.

- Create another switch case for all the given Subs/Functions. Do not add or remove any communication methods.

- Test the implementation. There are a few example methods for testing, feel free to add as necessary.

Imports and global variables can be added as needed.


The codebase can be run by creating a VB.NET command line project in Visual Studio and pasting/adding the code. You need to add ASCOM DLLs (see this video: https://www.youtube.com/watch?v=SfFg5xoVKhg&t=86) as a reference and should set the CPU to x64.
