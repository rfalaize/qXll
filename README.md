# qXll
kdb+ csharp wrapper for Microsoft Excel, using ExcelDna.
Leverage Excel capabilities using kdb+ in-memory data storage and vectorial computation.


More infos:
* kdb+: http://code.kx.com/q/
* ExcelDna: https://github.com/Excel-DNA/ExcelDna/


Versions: 
**1.0 - 2017-05-07 - Add qExecute, qQuery and qInsert**
* **qQuery**: run a q command and returns a 2D variant. Works with any command that returns a well formatted q table or keyed table.
* **qInsert**: insert a 2D variant in a q table. Creates it if necessary using the fist row as column names.
* **qExecute**: run any q command. Do not return any result. Can be used to create custom functions.
* All  functions have a synchronous and an asynchronous version.
      
**1.1 - 2017-07-02 - Add qSubscribe**
* **qSubscribe**: subscribe to real-time updates on a data point identified by a unique key.
      
      
