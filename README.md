# HorizontalVirtualizingWrapPanel

The native WrapPanel does not support Virtualization, which is a problem when displaying large datasets. This class only supports Horizontal, since that was all I needed. 

You can either set "ItemsPerRow"-dependency property to specify how many items each row should consist of, or you can just let the panel calculate how many items theres room for.

* THIS PANEL ONLY WORKS IF ALL CHILDREN ARE THE SAME SIZE.
