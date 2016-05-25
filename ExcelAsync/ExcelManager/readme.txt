All logics should be written by private method. Then private method will be exposed to public by QueueAsMacro. So we can
make sure all logic run in the excle main thread. 

We may need to operate excel multiple times according to business logic. We way pass in business model to this namespace.

Finally all excel operator should be in ExcelOperator namespace. 
