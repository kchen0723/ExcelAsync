# ExcelAsync
How to use Excel DNA to develop excel addin

Copied the Menu from http://brymck.github.io/finansu/install temporary for testing purpose. 

ExcelAsyncWpf should be renamed to ExcelAsync, which will operate Excel directly. All Excel operation should be in this project, such as Ribbon, custom panel, Formula entry, com server, context menu.
This project will call your custom logic by delegate.

ExcelAsyncWvvm is the project for custom logic. Utilize any .net technology here to implement your own logic. You can treat this project as a stand wpf application. If you need to operate excel, 
Expose the operation by delegate property. Then inject the delegate property by above project. 
This project should NOT reference the Microsoft office dll. All Excel operation should be done in above project.