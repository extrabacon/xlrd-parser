rm -rf deps
mkdir deps
cd deps
curl -s -L https://github.com/python-excel/xlrd/tarball/master | tar zx
mv python-excel-xlrd-* python-excel-xlrd
cd ..
cp runxlrd-json.py deps/python-excel-xlrd
