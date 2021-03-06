# qXll
kdb+ csharp wrapper for Microsoft Excel, using ExcelDna.
Leverage Excel capabilities using kdb+ in-memory data storage and vectorial computation.


### Versions 

#### 1.0 - 2017-05-07 - Add qExecute, qQuery and qInsert
* **qQuery**: run a q command and returns a 2D variant. Works with any command that returns a well formatted q table or keyed table.
* **qInsert**: insert a 2D variant in a q table. Creates it if necessary using the fist row as column names.
* **qExecute**: run any q command. Do not return any result. Can be used to create custom functions.
* All  functions have a synchronous and an asynchronous version.
      
#### 1.1 - 2017-07-02 - Add qSubscribe
* **qSubscribe**: subscribe to real-time updates on a data point identified by a unique key. Supports simultaneous connections to multiple q processes and servers in the same Excel instance.
     
#### 1.2 - 2017-07-07 - Add qProcessStart and qProcessKill
* **qProcessStart**: starts a q process on localhost. Supports command line parameters. If needed checks if port is available.
* **qProcessKill**: kills an existing q process on localhost.


### More infos
* kdb+: http://code.kx.com/q/
* ExcelDna: https://github.com/Excel-DNA/ExcelDna/
