# IEInProcessExecutor
Originally provided sample code with for a [specific StackOverflow question](https://stackoverflow.com/questions/52466540/how-can-i-create-a-javascript-object-in-internet-explorer-from-c).
This project was originally to give potential answerers a full-scale working prototype with which to confirm the behavior. Currently, the project now exists to demonstrate executing commands in-process within an instance of Internet Explorer.

The Visual Studio solution contained herein should be compilable using Visual Studio 2019. This includes the Community edition,
provided the Windows Desktop C++ workflow and support for Active Template Library (ATL) has been installed.

Note that most users will need to compile the solution using the "x86" platform to work correctly with Internet Explorer 11, even if the browser is installed and running on 64-bit Windows.
