Array.prototype.map = function (f) {
    for (var i = 0; i < this.length; i++) {
        f(this[i]);
    }
}

var fso = new ActiveXObject("Scripting.FileSystemObject");

var SearchOption = new (function SearchOption() {})();
SearchOption.constructor.name = "SearchOption";
SearchOption.constructor.prototype.TopDirectoryOnly = 0;
SearchOption.constructor.prototype.AllDirectories = 1;

var File = new (function File() { })();
File.constructor.name = "File";
//Directory.constructor.prototype.

var Directory = new (function Directory() { })();
Directory.constructor.name = "Directory";

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
            if (typeof item != "number") { return; }
//            if (!(item && item.constructor.name == "SearchOption")) { return; }
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
    var checkPattern = function (filename) {
        if (searchPattern) return true;
        if (searchPattern == "") return true;
        if (searchPattern == "*") return true;
        return (0 <= filename.indexOf(searchPattern));
    };
    var _getFiles = function (p) {
        var files = [];
        var f = fso.GetFolder(p);
        // ファイル
        var e = new Enumerator(f.files);
        for (; !e.atEnd(); e.moveNext()) {
            var file = e.item();
            if (checkPattern(file.Name)) {
                files.push(file.Path);
            }
        }
        // フォルダ
        var e = new Enumerator(f.SubFolders);
        for (; !e.atEnd(); e.moveNext()) {
            var file = e.item();
            if (searchOption == SearchOption.AllDirectories) {
                // 再帰
                var a = _getFiles(file.Path);
                if (a && a instanceof Array && 0 < a.length) {
                    files = files.concat(a);
                }
            }
        }
        return files;
    }
    return _getFiles(path);
};

function test001() {
    debugger;
    var files = Directory.GetFiles("D:\\develop\\github\\WSHSamples", "", SearchOption.AllDirectories);
//    WScript.Echo(files);
    files.map(function(file) {
        WScript.Echo(file);
    });
}
