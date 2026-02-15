class MergeCells:
    def __init__(self, raw_merges):
        self._merges = raw_merges

    def __len__(self):
        return self._merges.count()

    def __getitem__(self, index):
        return self._merges[index]

    def __iter__(self):
        for i in range(len(self)):
            yield self[i]

    def __contains__(self, item):
        return self._merges.merge_exists(item)

    def append(self, reference):
        """Append a new merged range (e.g. 'A1:B2')."""
        self._merges.append_merge(reference)

    def delete(self, index):
        """Delete a merged range by index."""
        self._merges.delete_merge(index)

    def find(self, reference):
        """Find index of a merged range. Returns -1 if not found."""
        return self._merges.find_merge(reference)
