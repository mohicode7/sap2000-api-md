# Item Type For Elements

The ItemTypeElm is one of the following items in the eItemTypeElmenumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the item request is for the point element corresponding to the point object specified by the Name item.

If the ItemTypeElm is Element, the item request is for the point element specified by the Name item.

If the ItemTypeElm is GroupElm, the item request is for all point elements directly or indirectly specified in the group specified by the Name item. If a point object is in the specified group, data is requested for its corresponding point element. If line, area, solid or link objects are in the specified group, data is requested for all point elements associated with those objects, including those point elements created as a result of automatic internal meshing of the object by the program.

If the ItemTypeElm is SelectionElm, the item request is for all point elements directly or indirectly selected and the Name item is ignored. If a point object is selected, data is requested for its corresponding point element. If line, area, solid or link objects are selected, data is requested for all point elements associated with those objects, including those point elements created as a result of automatic internal meshing of the object by the program.
