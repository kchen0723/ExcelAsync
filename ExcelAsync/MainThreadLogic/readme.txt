All excel operation should be written in this namespace. 
All classes in this namespace should be internal or private.
All internal method should call CheckIsExcelUIMainThread() method to make sure it run on Excel Main thread.

If a call is from context menu, formula, or ribbon, then it is on main thread.