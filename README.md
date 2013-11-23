photoshop_utils
===============

Python wrapper classes and utilities for controlling photoshop over win32com api.

This files contains simple wrapper classes for the api exposed by Adobe Photoshop via com interface. See http://www.adobe.com/devnet/photoshop/scripting.html for more.
It is still under construction and not fully tested. Though classes in this file are usually added to solve a task I face.

Right now it's only working for windows (win32com) but should be adaptable to mac. As soon as I get one, I'll start working on it. :)-

In addition to the wrapper classes there are some utility methods like recursively calling a function on all artlayers.

The classes were build for CS6.

Feel free to add more wrappers and utilities.
