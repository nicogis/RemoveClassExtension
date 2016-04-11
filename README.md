# RemoveClassExtension
Remove class extension and Instance class

# Description

This is an Add-in for ArcCatalog. In categories 'Remove class extension' you have two buttons. 

If you have a feature class that has a class extension associated or a feature class that is a custom object whose behavior class
you can use this add-in that remove the reference and set null.

When previewing a feature class or adding it to a map, one of the following error messages occurs: 
“Error opening feature class. 
Unable to create object class extension COM component.” 
“Error opening feature class. 
Unable to create object class COM component.” 

In the first case, the feature class has a class extension associated that is not installed on the client machine. 
Using remove class extension you resolve this problem. 

In the second case, the feature class is a custom object whose behavior class is not installed on the client machine. 
Using remove instance class you resolve this problem.

This add-in runs also for annotation and dimension with class extension or custom object

# Requirements 
ArcGIS 10.4 Desktop or superior
