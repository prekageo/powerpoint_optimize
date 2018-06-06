#!/usr/bin/env python3

import multiprocessing
import os
import subprocess
import sys
import zipfile

class File:
    def __init__(self, zip, path):
        if not isinstance(zip, zipfile.ZipFile):
            zip = zipfile.ZipFile(zip)
        self.data = zip.open(path).read()
        self.path = path

    def is_same(self, path):
        return self.path == path

    @classmethod
    def process(cls, *args):
        obj = cls(*args)
        obj._process()
        return obj

class PNG(File):
    @staticmethod
    def matches(path):
        path = path.lower()
        if not path.startswith('ppt/media/'):
            return False
        if not path.endswith('.png'):
            return False
        return True

class PNG_Optipng(PNG):
    def _process(self):
        tmpname = '/tmp/%s' % os.path.basename(self.path)
        with open(tmpname, 'wb') as f:
            f.write(self.data)
        subprocess.check_call(['optipng', '-clobber', '-strip', 'all', tmpname], stderr=subprocess.DEVNULL)
        self.data = open(tmpname, 'rb').read()
        self.pathout = self.path

class PNG_PNGToJPG(PNG):
    def _process(self):
        tmpname = '/tmp/%s' % os.path.basename(self.path)
        with open(tmpname, 'wb') as f:
            f.write(self.data)
        jpgname = PNG_PNGToJPG.to_jpg(tmpname)
        subprocess.check_call(['convert', tmpname, '-background', 'white', '-flatten', '-alpha', 'off', jpgname])
        jpg_size = os.path.getsize(jpgname)
        if jpg_size < len(self.data):
            self.data = open(jpgname, 'rb').read()
            self.pathout = PNG_PNGToJPG.to_jpg(self.path)
        else:
            self.path = None
            self.data = None

    @staticmethod
    def to_jpg(path):
        return os.path.splitext(path)[0] + '.jpg'

class REL(File):
    def __init__(self, zip, path, pngs):
        super().__init__(zip, path)
        self.pngs = pngs

    def _process(self):
        out = self.data
        for png in self.pngs:
            if png.path is None:
                continue
            out = out.replace(os.path.basename(png.path).encode('ascii'), os.path.basename(png.pathout).encode('ascii'))
        if out == self.data:
            self.path = None
            self.data = None
        else:
            self.data = out
            self.pathout = self.path

    @staticmethod
    def matches(path):
        return path.lower().startswith('ppt/slides/_rels/')

def process(inputname, inputzip, klass, extra_args = None, pool = None):
    if not extra_args:
        extra_args = []
    results = []
    for path in inputzip.namelist():
        if klass.matches(path):
            if pool:
                results.append(pool.apply_async(klass.process, (inputname, path, *extra_args)))
            else:
                results.append(klass.process(inputzip, path, *extra_args))
    if pool:
        results = [r.get() for r in results]
    return results

def write_output(inputzip, outputzip, results):
    for info in inputzip.infolist():
        for r in results:
            if r.is_same(info.filename):
                info.compress_type = zipfile.ZIP_DEFLATED
                info.filename = r.pathout
                outputzip.writestr(info, r.data)
                break
        else:
            data = inputzip.open(info.filename).read()
            info.compress_type = zipfile.ZIP_DEFLATED
            outputzip.writestr(info, data)

def main():
    mode = sys.argv[1]
    input = sys.argv[2]
    output = sys.argv[3]

    pool = multiprocessing.Pool(32)
    inputzip = zipfile.ZipFile(input)
    outputzip = zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED)

    if mode == '--optipng':
        print('Starting optipng...')
        results_png = process(input, inputzip, PNG_Optipng, pool = pool)
        print('Finished optipng')
        write_output(inputzip, outputzip, results_png)
    elif mode == '--png-to-jpg':
        results_png = process(input, inputzip, PNG_PNGToJPG, pool = pool)
        results_rel = process(input, inputzip, REL, (results_png,))
        write_output(inputzip, outputzip, results_png + results_rel)
    else:
        sys.exit(1)

    outputzip.close()

main()
