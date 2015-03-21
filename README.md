# SolidworksUtils
Various utilities for Solidworks

Snapshotter
-----------

Automatically generates step files and jpegs of parts and assemblies and pdfs of drawings.

smgImport, smgExport
-----------

Text-based user-centric description of parts. Can be used (or could be, when finished) as a git merge tool for solidworks. Represents *all* user input that was used to generate the part in a common yaml format, with the goal of allowing native model history tree communication between different CAD packages. Can also be used to provide backwards compatibility to different versions of solidworks (2014 -> 2013, etc), so long as the basic construction of the objects remains unchanged.

As of 20 March 2015, export is passable (MVP for sure), but import is severely lacking. No assembly or drawing support exists yet, only parts.