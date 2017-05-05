# VBA Project: FadoSorting

Collection of routines to process special GSC document for alignment and final doc production:

- Master_Process_Doc_For_Alignment - Run this to prepare doc by removing and converting heavy stuff, not needed for aligment
- Master_Sort_ActiveDoc_FadoGlossary (incomplete - does not do preparations/ prior verifications) - Run this to sort current doc acc to linguistic order of its chapters titles, which it extracts by itself
- Master_Sort_TargetDoc_forAlignment_FadoGlossary (almost complete) - Run this with source-target languages documents opened to auto-extract the chapter id order used in source and sort target doc accordingly

    - one issue with this: misses one check, to see whether all chapters in both documents produce exactly the same main IDs (those are identified by one hard space before ### ID number, secondaries with simple space before. We're using this, so if this is not consistent, no good results! Perhaps to fix)






This repo (FadoSorting) was automatically created on 04/05/2017 19:26:06 by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/vbagit "desktop liberation")
you can see [library and dependency information here](dependencies.md)

To get started with VBA Git, you can either create a Document with the [code on gitHub](https://github.com/brucemcpherson/VbaGit "VbaGit repo"), or use this premade [VbaBootStrap Document](http://ramblings.mcpher.com/Home/excelquirks/downlable-items/VbaGitBootStrap.xlsm "VbaBootStrap")  


