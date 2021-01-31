import os
from datetime import datetime
from typing import Any, Callable


class FileBasic(object) :
    def __init__(self, path) :
        if not os.path.isfile(path) :
            raise FileNotFoundError(
                f"해당 파일이 존재하지 않거나 올바른 파일 확장자가 아닙니다. \n    >> {path}")
        self._absPath = os.path.abspath(path)
        self._fileName = os.path.basename(path)
        self._ext = os.path.splitext(self._fileName)[1]

        getDatetime: Callable[[Any],
                              datetime] = lambda x :datetime.fromtimestamp(x)
        self._createdTime = getDatetime(os.path.getctime(path))
        self._modifiedTime = getDatetime(os.path.getmtime(path))
        self._accessedTime = getDatetime(os.path.getatime(path))
        self._size = os.path.getsize(path)

    @property
    def path(self) :
        return self._absPath

    @property
    def fileName(self) :
        return self._fileName

    @property
    def folderName(self) :
        return os.path.basename(os.path.dirname(self.path))

    @property
    def ext(self) :
        return self._ext

    @property
    def createdTime(self) :
        return self._createdTime

    @property
    def modifiedTime(self) :
        return self._modifiedTime

    @property
    def accessedTime(self) :
        return self._accessedTime

    @property
    def size(self) :
        return self._size

    def isModified(self) :
        mtime = datetime.fromtimestamp(os.path.getmtime(self._absPath))
        if mtime != self._modifiedTime :
            return True
        else :
            return False

    def __str__(self) :
        timeFormat = "%Y-%m-%d %H:%M:%S"
        return f"absPath > {self._absPath}\n" \
            f"fileName > {self._fileName}\n" \
            f"ext > {self._ext}\n" \
            f"c_time > {self._createdTime.strftime(timeFormat)}\n" \
            f"m_time > {self._modifiedTime.strftime(timeFormat)}\n" \
            f"a_time > {self._accessedTime.strftime(timeFormat)}\n" \
            f"size > {self._size} byte"
