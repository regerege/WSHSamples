var fso = new ActiveXObject("Scripting.FileSystemObject");

var SearchOption = new Object();
SearchOption.constructor.prototype.TopDirectoryOnly = 0;
SearchOption.constructor.prototype.AllDirectories = 1;


var Directory = new Object();
//Directory.constructor._Init = function () { Directory.constructor._fso = new ActiveXObject("Scripting.FileSystemObject"); }
//Directory.constructor._Dispose = function () { this._fso = undefined; }

/// <param name="path">検索するディレクトリ</param>
/// <param name="searchPattern">
/// path 内のファイル名と対応させる検索文字列。
/// このパラメータは、2 つのピリオド ("..") で終了することはできません。
/// また、2 つのピリオド ("..") に続けて DirectorySeparatorChar または AltDirectorySeparatorChar を指定したり、InvalidPathChars の文字を含めたりすることはできません。
/// </param>
/// <param name="searchOption">
/// 検索操作にすべてのサブディレクトリを含めるのか、または現在のディレクトリのみを含めるのかを指定する SearchOption 値の 1 つ。
/// </param>
Directory.constructor.prototype.GetFiles = function () {
    var len = arguments.length;
    var path = "";
    var searchPattern = "*";
    var searchOption = SearchOption.TopDirectoryOnly;
    switch (arguments.length) {
        case 3:
            var item = arguments[2];
            if (!(item instanceof SearchOption)) { return; }
            searchOption = item;
        case 2:
            var item = arguments[1];
            if (typeof item != "string") { return; }
            searchPattern = item;
        case 1:
            var item = arguments[0];
            if (typeof item != "string") { return; }
            path = item;
            break;
        default:
            return;
    }
    var _getFiles = function(p) {
        var files = [];
        var f = fso.GetFolder(p);
        var e = new Enumerator(f.files);
        for (; !e.atEnd(); e.moveNext()) {
            var file = e.item();
            if (0 <= file.indexOf(searchPattern) || searchPattern == "*") {
                files.push(file);
            }
        }
        var e = new Enumerator(f.SubFolders);
        for (; !e.atEnd(); e.moveNext()) {
            var file = e.item();
            if (0 <= file.indexOf(searchPattern)) {
                files.push(file);
            }
            // 再帰
            if (searchOption == SearchOption.AllDirectories) {
                var a = _getFiles(p);
                if (a && a instanceof Array && 0 < a.length) {
                    files.concat(a);
                }
            }
        }
        return files;
    }
    return _getFiles(p);
};

//if (!fso) {
//    var fso = new ActiveXObject("Scripting.FileSystemObject");
//}
var files = Directory.GetFiles("D:\\develop\\github");
var a = fso.CreateTextFile("D:\\testfile.txt", true);
a.WriteLine("ファイルリスト");
for(var file in files) {
    a.WriteLine(file);
}
a.Close();
