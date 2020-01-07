# Gassample
Implementation example of Google Apps Script with clasp

## 前提条件

以下のリソースをenableにすること。
- Drive API v2
- Docs API v1

## 動作環境

mac OS Catalina 
version 10.15.2

```
$ sw_vers
ProductName:	Mac OS X
ProductVersion:	10.15.2
BuildVersion:	19C57
$ npm --version
6.13.4
$ ../node_modules/.bin/clasp --version
2.3.0
```
(※claspはnpm install -gしてません。)

## Claspの操作手順(覚えられないので)

### pull

```
$ ../node_modules/.bin/clasp pull
```

### push (Note: https://script.google.com/home/usersettings でAPIを有効にする必要がある。)

```
$ ../node_modules/.bin/clasp push
└─ appsscript.json
└─ code.js
Pushed 2 files.
```