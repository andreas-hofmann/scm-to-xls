#!/usr/bin/env python3

from pygit2 import Repository
from pygit2 import GIT_SORT_TOPOLOGICAL, GIT_SORT_REVERSE

from openpyxl import Workbook

class ScmAccessor:
    def __init__(self):
        raise NotImplementedError("Implement me!")


class GitAccessor(ScmAccessor):
    def __init__(self):
        raise NotImplementedError("Implement me!")


class HgAccessor(ScmAccessor):
    def __init__(self):
        raise NotImplementedError("Implement me!")


class SvnAccessor(ScmAccessor):
    def __init__(self):
        raise NotImplementedError("Implement me!")


if __name__ == "__main__":
    print("Nothing here yet :(")
