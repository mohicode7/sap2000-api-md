# RefreshWindow and RefreshView

The RefreshWindow and RefreshView functions are similar. RefreshWindow simply updates the display. RefreshView rebuilds all object arrays used for display and then updates the display. A RefreshView takes longer to complete than a RefreshWindow.

The RefreshView function should be called after operations that add, delete or move objects. The RefreshWindow function is adequate for other types of changes. For example, after adding a new point object to the model, the  RefreshView function should be called. After modifying a joint restraint, the RefreshWindow function should be called.
