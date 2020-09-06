#include <File.au3>
#include "..\Class.au3"

$aFiles = _FileListToArrayRec(@ScriptDir, '*.au3p', $FLTA_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
For $i = 1 To $aFiles[0] Step +1
    If FileExists(StringTrimRight($aFiles[$i], 1)) Then FileDelete(StringTrimRight($aFiles[$i], 1))
    Class_Convert_File($aFiles[$i])
Next
