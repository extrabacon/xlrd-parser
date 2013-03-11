# <p>Copyright (c) 2013 Nicolas Mercier</p>
# <p>This script has been modified from the official xlrd package to output JSON</p>

# <p>Copyright (c) 2005-2012 Stephen John Machin, Lingfo Pty Ltd</p>
# <p>This script is originally part of the xlrd package, which is released under a
# BSD-style licence.</p>

from __future__ import print_function

cmd_doc = """
Commands:

2rows           Print the contents of first and last row in each sheet
3rows           Print the contents of first, second and last row in each sheet
bench           Same as "show", but doesn't print -- for profiling
hdr             Mini-overview of file (no per-sheet information)
hotshot         Do a hotshot profile run e.g. ... -f1 hotshot bench bigfile*.xls
ov              Overview of file
profile         Like "hotshot", but uses cProfile
show            Print the contents of all rows in each sheet
version[0]      Print versions of xlrd and Python and exit

[0] means no file arg
[1] means only one file arg i.e. no glob.glob pattern
"""

options = None
if __name__ == "__main__":

    PSYCO = 0

    import xlrd
    import sys, time, glob, traceback, pprint, gc, json
    
    from xlrd.timemachine import xrange, REPR
    

    class LogHandler(object):

        def __init__(self, logfileobj):
            self.logfileobj = logfileobj
            self.fileheading = None
            self.shown = 0
            
        def setfileheading(self, fileheading):
            self.fileheading = fileheading
            self.shown = 0
            
        def write(self, text):
            if self.fileheading and not self.shown:
                self.logfileobj.write(self.fileheading)
                self.shown = 1
            self.logfileobj.write(text)
        
    null_cell = xlrd.empty_cell

    def show_row(bk, sh, rowx, colrange, printit):
        if bk.ragged_rows:
            colrange = range(sh.row_len(rowx))
        if not colrange: return
        for colx, ty, val, cxfx in get_row_data(bk, sh, rowx, colrange):
            if printit:
                print(json.dumps([ "cell", { "r": rowx, "c": colx, "cn": xlrd.colname(colx), "t": ty, "v": val }]))

    def get_row_data(bk, sh, rowx, colrange):
        result = []
        dmode = bk.datemode
        ctys = sh.row_types(rowx)
        cvals = sh.row_values(rowx)
        for colx in colrange:
            cty = ctys[colx]
            cval = cvals[colx]
            if bk.formatting_info:
                cxfx = str(sh.cell_xf_index(rowx, colx))
            else:
                cxfx = ''
            if cty == xlrd.XL_CELL_DATE:
                try:
                    showval = xlrd.xldate_as_tuple(cval, dmode)
                except xlrd.XLDateError:
                    e1, e2 = sys.exc_info()[:2]
                    showval = "%s:%s" % (e1.__name__, e2)
                    cty = xlrd.XL_CELL_ERROR
            elif cty == xlrd.XL_CELL_ERROR:
                showval = xlrd.error_text_from_code.get(cval, '<Unknown error code 0x%02x>' % cval)
            else:
                showval = cval
            result.append((colx, cty, showval, cxfx))
        return result

    def show(bk, nshow=65535, printit=1):
        if options.onesheet:
            try:
                shx = int(options.onesheet)
            except ValueError:
                shx = bk.sheet_by_name(options.onesheet).number
            shxrange = [shx]
        else:
            shxrange = range(bk.nsheets)
        for shx in shxrange:
            sh = bk.sheet_by_index(shx)
            nrows, ncols = sh.nrows, sh.ncols
            colrange = range(ncols)
            print(json.dumps([ "sheet", {
                "index": shx,
                "name": sh.name,
                "rows": sh.nrows,
                "cols": sh.ncols,
                "visibility": sh.visibility
            }]))
            if nrows and ncols:
                # Beat the bounds
                for rowx in xrange(nrows):
                    nc = sh.row_len(rowx)
                    if nc:
                        _junk = sh.row_types(rowx)[nc-1]
                        _junk = sh.row_values(rowx)[nc-1]
                        _junk = sh.cell(rowx, nc-1)
            for rowx in xrange(nrows-1):
                if not printit and rowx % 10000 == 1 and rowx > 1:
                    print("done %d rows" % (rowx-1,))
                show_row(bk, sh, rowx, colrange, printit)
            if nrows:
                show_row(bk, sh, nrows-1, colrange, printit)
            if bk.on_demand: bk.unload_sheet(shx)

    def main(cmd_args):
        import optparse
        global options, PSYCO
        usage = "\n%prog [options] command [input-file-patterns]\n" + cmd_doc
        oparser = optparse.OptionParser(usage)
        oparser.add_option(
            "-l", "--logfilename",
            default="",
            help="contains error messages")
        oparser.add_option(
            "-v", "--verbosity",
            type="int", default=0,
            help="level of information and diagnostics provided")
        oparser.add_option(
            "-p", "--pickleable",
            type="int", default=1,
            help="1: ensure Book object is pickleable (default); 0: don't bother")
        oparser.add_option(
            "-m", "--mmap",
            type="int", default=-1,
            help="1: use mmap; 0: don't use mmap; -1: accept heuristic")
        oparser.add_option(
            "-e", "--encoding",
            default="",
            help="encoding override")
        oparser.add_option(
            "-f", "--formatting",
            type="int", default=0,
            help="0 (default): no fmt info\n"
                 "1: fmt info (all cells)\n"
            )
        oparser.add_option(
            "-g", "--gc",
            type="int", default=0,
            help="0: auto gc enabled; 1: auto gc disabled, manual collect after each file; 2: no gc")
        oparser.add_option(
            "-s", "--onesheet",
            default="",
            help="restrict output to this sheet (name or index)")
        oparser.add_option(
            "-u", "--unnumbered",
            action="store_true", default=0,
            help="omit line numbers or offsets in biff_dump")
        oparser.add_option(
            "-d", "--on-demand",
            action="store_true", default=0,
            help="load sheets on demand instead of all at once")
        oparser.add_option(
            "-r", "--ragged-rows",
            action="store_true", default=0,
            help="open_workbook(..., ragged_rows=True)")
        options, args = oparser.parse_args(cmd_args)
        if len(args) == 1 and args[0] in ("version", ):
            pass
        elif len(args) < 2:
            oparser.error("Expected at least 2 args, found %d" % len(args))
        cmd = args[0]
        xlrd_version = getattr(xlrd, "__VERSION__", "unknown; before 0.5")
        if cmd == 'biff_dump':
            xlrd.dump(args[1], unnumbered=options.unnumbered)
            sys.exit(0)
        if cmd == 'biff_count':
            xlrd.count_records(args[1])
            sys.exit(0)
        if cmd == 'version':
            print("xlrd: %s, from %s" % (xlrd_version, xlrd.__file__))
            print("Python:", sys.version)
            sys.exit(0)
        if options.logfilename:
            logfile = LogHandler(open(options.logfilename, 'w'))
        else:
            logfile = sys.stdout
        mmap_opt = options.mmap
        mmap_arg = xlrd.USE_MMAP
        if mmap_opt in (1, 0):
            mmap_arg = mmap_opt
        elif mmap_opt != -1:
            print("Unexpected value (%r) for mmap option -- assuming default" % mmap_opt)
        fmt_opt = options.formatting | (cmd in ('xfc', ))
        gc_mode = options.gc
        if gc_mode:
            gc.disable()
        for pattern in args[1:]:
            for fname in glob.glob(pattern):
                if logfile != sys.stdout:
                    logfile.setfileheading("\n=== File: %s ===\n" % fname)
                if gc_mode == 1:
                    n_unreachable = gc.collect()
                    if n_unreachable:
                        print("GC before open:", n_unreachable, "unreachable objects")
                if PSYCO:
                    import psyco
                    psyco.full()
                    PSYCO = 0
                try:
                    t0 = time.time()
                    bk = xlrd.open_workbook(fname,
                        verbosity=options.verbosity, logfile=logfile,
                        pickleable=options.pickleable, use_mmap=mmap_arg,
                        encoding_override=options.encoding,
                        formatting_info=fmt_opt,
                        on_demand=options.on_demand,
                        ragged_rows=options.ragged_rows,
                        )
                    t1 = time.time()
                    print(json.dumps(["workbook", { "file": fname, "user": bk.user_name, "sheets": { 
                        "count": bk.nsheets,
                        "names": bk.sheet_names()
                    }}]))
                except xlrd.XLRDError:
                    e0, e1 = sys.exc_info()[:2]
                    print(json.dumps(["error", {
                        "type": "open_failed",
                        "exception": e0.__name__,
                        "message": e1.message
                    }]))
                    continue
                except KeyboardInterrupt:
                    print(json.dumps(["error", { "type": "keyboard_interrupt", "message": "keyboard interrupt" }]))
                    #traceback.print_exc(file=sys.stdout)
                    sys.exit(1)
                except:
                    e0, e1 = sys.exc_info()[:2]
                    print(json.dumps(["error", {
                        "type": "open_failed",
                        "exception": e0.__name__,
                        "message": e1.message
                    }]))
                    #traceback.print_exc(file=sys.stdout)
                    continue
                t0 = time.time()
                if cmd == 'ov': # OverView
                    show(bk, 0)
                elif cmd == 'show': # all rows
                    show(bk)
                elif cmd == '2rows': # first row and last row
                    show(bk, 2)
                elif cmd == '3rows': # first row, 2nd row and last row
                    show(bk, 3)
                elif cmd == 'bench':
                    show(bk, printit=0)
                else:
                    print(json.dumps(["error", { "type": "unknown_command", "message": "Unknown command: %s" % cmd }]))
                    sys.exit(1)
                del bk
                if gc_mode == 1:
                    n_unreachable = gc.collect()
                    if n_unreachable:
                        print("GC post cmd:", fname, "->", n_unreachable, "unreachable objects")

        return None

    av = sys.argv[1:]
    if not av:
        main(av)
    firstarg = av[0].lower()
    if firstarg == "hotshot":
        import hotshot, hotshot.stats
        av = av[1:]
        prof_log_name = "XXXX.prof"
        prof = hotshot.Profile(prof_log_name)
        # benchtime, result = prof.runcall(main, *av)
        result = prof.runcall(main, *(av, ))
        print("result", repr(result))
        prof.close()
        stats = hotshot.stats.load(prof_log_name)
        stats.strip_dirs()
        stats.sort_stats('time', 'calls')
        stats.print_stats(20)
    elif firstarg == "profile":
        import cProfile
        av = av[1:]
        cProfile.run('main(av)', 'YYYY.prof')
        import pstats
        p = pstats.Stats('YYYY.prof')
        p.strip_dirs().sort_stats('cumulative').print_stats(30)
    elif firstarg == "psyco":
        PSYCO = 1
        main(av[1:])
    else:
        main(av)
