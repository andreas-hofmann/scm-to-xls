#!/usr/bin/env python3

from os import path, getcwd
from sys import exit

from pygit2 import Repository
from pygit2 import GIT_SORT_TIME, GIT_DIFF_STATS_FULL

from openpyxl import Workbook

from optparse import OptionParser
from datetime import datetime

class LogEntry:
    def __init__(self):
        self.id = None
        self.msg = None
        self.author = None
        self.email = None
        self.time = None
        self.diff = None

class ScmAccessor:
    def __init__(self, repo_path):
        self._repo_path = repo_path
        self._scm = None

    def get_log(self):
        raise RuntimeError("Not meant to be called in parent class")


class GitAccessor(ScmAccessor):
    def __init__(self, repo_path):
        super().__init__(repo_path)
        self._scm = Repository(path.join(repo_path))

    def get_log(self):
        data = []
        for c in self._scm.walk(self._scm.head.target, GIT_SORT_TIME):
            diff = self._scm.diff(c, c.parents[0]).stats.format(GIT_DIFF_STATS_FULL, 1) if c.parents else ""

            diff = diff.splitlines()
            if len(diff) >= 1:
                diff = diff[:-1]

            stripped_diff = [ d.split("|")[0].strip() for d in diff ]

            e = LogEntry()
            e.id = c.id
            e.msg = c.message
            e.author = c.committer.name
            e.email = c.committer.email
            e.time = c.commit_time
            e.diff = stripped_diff
            data.append(e)
        return data


class HgAccessor(ScmAccessor):
    def __init__(self, repo_path):
        super().__init__(repo_path)
        raise NotImplementedError("Implement me!")


class SvnAccessor(ScmAccessor):
    def __init__(self, repo_path):
        super().__init__(repo_path)
        raise NotImplementedError("Implement me!")



class Writer:
    def __init__(self, filename):
        self._filename = filename
        self._workbook = Workbook()

    def write_header(self):
        raise RuntimeError("Not to be called in base class")

    def write_data(self):
        raise RuntimeError("Not to be called in base class")

    def write_data(self, accessor):
        ws = self._workbook.active

        for d in accessor.get_log():
            ws.append([str(d.id), str(d.msg), str(d.author), str(d.email), str(d.time), "\n".join(d.diff)])

    def save(self):
        self._workbook.save(self._filename)


class CommitHistoryWriter(Writer):
    def __init__(self, filename):
        super().__init__(filename)

    def write_header(self):
        ws = self._workbook.active
        ws['A1'] = "Commit history"
        ws['B1'] = "Generated on %s" % datetime.now().isoformat(" ")
        ws.append(["Commit-ID", "Message", "Author", "E-Mail", "Date", "Changed files"])


class ImpactStatementWriter(Writer):
    def __init__(self, filename):
        super().__init__(filename=filename)

    def write_header(self):
        ws = self._workbook.active
        ws['A1'] = "Impact statement"
        ws['B1'] = "Generated on %s" % datetime.now().isoformat(" ")
        ws.append(["Commit-ID", "Message", "Author", "E-Mail", "Date", "Changed files", "Affected testcases", "Tested with version"])


def main():
    parser = OptionParser()

    parser.add_option("-o", "--outfile", dest="outfile",
                      help="Output file to write to.")
    parser.add_option("-s", "--scm", dest="scm",
                      help="SCM to use. Supported options: git, hg, svn",
                      default="git")
    parser.add_option("-I", "--impact", dest="impact",
                      help="If set, an impact statement is generated.",
                      action="store_true")
    parser.add_option("-H", "--history", dest="history",
                      help="If set, the commit history is exported.",
                      action="store_true")
    parser.add_option("-r", "--repo", dest="repo",
                      help="Path to the repository. If omitted, the location "
                            "the script is placed at is used.")

    options, args = parser.parse_args()

    if not options.history and not options.impact:
        print("Not sure what to do. Impact analysis (-I) or commit history (-H)?")
        exit(-1)

    if not options.outfile:
        print("Output filename missing.")
        exit(-1)

    rpath = options.repo if options.repo else getcwd()

    accessor = None

    if options.scm == "git":
        accessor = GitAccessor(rpath)
    elif options.scm == "hg":
        accessor = HgAccessor(rpath)
    elif options.scm == "svn":
        accessor = SvnAccessor(rpath)

    outfile = options.outfile

    if not outfile.endswith(".xlsx"):
        outfile += ".xlsx"

    if options.history:
        print("Generating commit history...")
        writer = CommitHistoryWriter("History-" + outfile)
    elif options.impact:
        print("Generating impact statement...")
        writer = ImpactStatementWriter("Impacts-" + outfile)

    writer.write_header()
    writer.write_data(accessor=accessor)
    writer.save()

    print("...done.")

if __name__ == "__main__":
    main()
