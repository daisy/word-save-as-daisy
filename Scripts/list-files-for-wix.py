from os import listdir, remove
from os.path import isfile, join, dirname, basename, exists
import glob
import uuid
import re

_oldroot= join(dirname(dirname(__file__)), "Lib")
_newroot="$(var.SolutionDir)Lib"
_indent = "    "
_mediaIdStart = 2
pathes = ['daisy-pipeline']


def print_wix_tree(path, _mediaId, level=0):
    '''Recursively parse a folder to build a wix project directory tree
    path [string] : path of the directory to parse
    level [int=0] : starting indentation level'''
    content = ''
    refs = ''
    if isfile(path):
        content = (_indent * level) +\
            "<File DiskId=\"%s\" Id=\"%s\" Name=\"%s\" Source=\"%s\" />\n" %(
                _mediaId,
                re.sub(r'\\|-|\$| ',"_",path[len(_oldroot):]),
                basename(path),
                _newroot + path[len(_oldroot):]
                )
        return (content, refs)
    else: # path is a directory
        content = (_indent * level) +\
            "<Directory Id=\"%s\" Name=\"%s\">\n" %(
                re.sub(r'\\|-|\$',"_",path[len(_oldroot):]),
                basename(path))
        subpathes = listdir(path)
        files = []
        directories = []
        for subpath in subpathes:
            if isfile(join(path,subpath)):
                files.append(join(path,subpath))
            else:
                directories.append(join(path,subpath))
        if len(files)>0:
            id = re.sub(r'\\|-|\$',"_",path[len(_oldroot):]) + "_files"
            content += (_indent * (level + 1)) +\
                "<Component Id=\"%s\" Guid=\"%s\">\n" %(id, str(uuid.uuid4()))
            refs += (_indent * (level + 1)) + "<ComponentRef Id=\"%s\"/>\n" %(id)
            for file in files:
                (_cont, _refs) = print_wix_tree(file, _mediaId, level + 2)
                content += _cont
                refs += _refs
            content += (_indent * (level + 1)) + "</Component>\n"
        for directory in directories:
            (_cont, _refs) = print_wix_tree(directory, _mediaId, level + 1)
            content += _cont
            refs += _refs
        content += (_indent * level) + "</Directory><!--%s-->\n"%(basename(path))
        return (content, refs)



for idx, _path in enumerate(pathes):
    print(join(_oldroot, _path))
    (cont, refs) = print_wix_tree(join(_oldroot, _path), _mediaIdStart + idx, 6)
    if exists(_path + "-dirs.wix"):
        remove(_path + "-dirs.wix")
    with open(_path + "-dirs.wix" ,"a") as f:
        print(cont, file=f)
        print(refs, file=f)

# with open("output.txt" ,"a") as f:
#     print(print_wix_tree(join(_oldroot, _path),6), file=f)

