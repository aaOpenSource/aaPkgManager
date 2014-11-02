
aaPkgManager
------------

**What Is it?**
The aaPkgManager is a dll that can be used to manipulate System Platform aapkg files.

**What does it do?**

 - We can unpack all contents of an aapkg into a directory
 - We can repack all contents into an aapkg that can be imported by System Platform
 - We can remove aaPDF files from the package.  This saves a TON of space.
 - We can remove all instances and/or templates
 - We can specify which objects to leave in a package by name, removing everything else.

**What it doesn't do.. yet**

 - Create a perfect package that can import with no annoying errors
 - Extract the manifest file
 - Inject a manifest file

**What's on the Roadmap?**

 - Figure out how to create a package that imports with no extraneous errors
 - Add methods to allow extraction and insertion of manifest data
 - Write a tool that will create an index of the contents of multiple aapkg files.  There a a lot of potential opportunities with this, one of which might be to combine this with [aaExport](https://github.com/aaOpenSource/aaExport) and create an open source version control system for System Platform.
