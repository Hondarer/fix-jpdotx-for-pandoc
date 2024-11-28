# fix-jpdotx-for-pandoc

Pandoc で .docx ファイルを生成するために必要なテンプレートを作成する際、日本語版の Word で保存すると、いくつかの スタイルID が変更されてしない、Pandoc で認識されなくなる。
(例: Title というスタイル名を日本語版の Word で保存すると、スタイル名が 表題 になり、あわせて スタイルID が Title ではなくなるため、Pandoc で認識されなくなる。)

このコマンドは、日本語版の Word で保存したスタイルの スタイルID を Pandoc で認識可能な スタイルID にパッチする。

Pandoc のスタイル定義は、[https://github.com/jgm/pandoc/blob/main/data/docx/word/styles.xml](https://github.com/jgm/pandoc/blob/main/data/docx/word/styles.xml) にある。
