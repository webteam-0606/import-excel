## excel 自动导入数据库步骤
> 利用 Postgres 数据库，MySQL数据库，Navicat 16 工具，VScode工具。

1、进入文件夹，打开对应项目。
> 启动 `vscode`，文件 -> 打开文件夹 -> 选择 `import-excel` 文件夹。

+ 注意：如需上传的excel表名与数据库已有表重名，请到 步骤4，删除旧的重复的表。再进行一下操作。

2、修改 `excel` 路径，以及对应数据库名称。
> 点击 `excelTodatabase.py`， 修改以下内容：
```java
// 其中path为需要导入数据库的excel的路径。
// database 为导入excel的数据库名称。
def createtable():
    path='C:/Users/320200255/project/import-excel/test.xlsx'
    ...
    conn = psycopg2.connect(database='postgres-uploadsearch', user='postgres', password='postgres', host='localhost')
    ...
```
3、运行项目。
> 修改好，点击保存。右键 `Run code`，运行。

4、打开 `postgreSQL` 数据库，刷新表。
> 打开电脑应用程序 `postgreSQL 11` 中的 `pgAdmin 4` 应用，账户密码都为 `postgres`。依次点击 `postgres-uploadsearch` -> `postgres-uploadsearch` -> `public` -> 表。右键 `表`, 进行刷新，即可看到 `excel` 中的每个 `sheet` 表。

5、将 `postgreSQL` 中的表复制到 `MySql`。
> 将 `postgreSQL` 中的表复制到 `MySql`。本地 `excel` 导入 `postgreSQL` 数据库完成。
---------------------------
## vue 实现多表查找搜索
1、进入文件夹，打开对应项目。
> 启动 `vscode`，文件 -> 打开文件夹 -> 选择 `vue-excel-process-system-master` 文件夹。

2、修改 server API。index.js。router.js。
> ...

3、启动服务器。
> 终端输入 cd server。终端输入 node app.js。若显示 服务器启动成功 则ok。

4、启动项目。
> 终端界面，点击 powershell 右边的 + 后，终端再输入 npm run dev 启动项目。网页输入生成的地址，即可。

5、打包项目。
> 终端输入 ctrl+c 停止项目，终端输入 npm run build 打包项目。此时项目目录会自动更新dist文件夹。将dist文件夹复制到内网，替换原先的 `dist` 文件夹。

6、内网访问项目。
> 依次点击 `public` 文件夹 -> `dist` 文件夹 -> `index.html`。即可打开网页，进行搜索。

