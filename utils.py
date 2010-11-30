# coding=utf-8

import csv, re, xlrd


class KSHReader(object):
    u"""Reader class for KSH workbook."""
    DATA_SHEET_NAME = u'Minden helység adata'
    COL_PLACE = 0
    COL_KSHCODE = 1
    ROW_DATA_START = 3

    def __init__(self, filename):
        workbook = xlrd.open_workbook(filename)
        self.sheet = workbook.sheet_by_name(KSHReader.DATA_SHEET_NAME)

    def items(self):
        u"""Return a tuple that consists tuples of place name and KSH codes."""
        data = []
        for row in range(self.sheet.nrows):
            if row < KSHReader.ROW_DATA_START: continue # skip headings
            place, kshcode = self.sheet.row_values(row, start_colx=KSHReader.COL_PLACE,
                                                   end_colx=KSHReader.COL_KSHCODE + 1)
            if not place: continue # skip empty rows

            # Place value 'Budapest' is a special case, it is written as
            # 'Budapest X. kerület' in the data source.  Need to convert.
            if re.search('Budapest', place):
                place = u'Budapest'

            data.append((place, kshcode))
        return data


class GeoNamesReader(object):
    u"""Reader class for GeoNames postcode data file."""
    COL_POSTCODE = 1
    COL_PLACE = 2
    COL_LAT = 9
    COL_LNG = 10

    def __init__(self, filename):
        file = open(filename, 'r')
        self.reader = csv.reader(file, delimiter="\t", quotechar=None)

    def items(self):
        u"""Return a tuple that consists tuples of place name, post code,
        latitude, and longitude.
        """
        data = []
        for row in self.reader:
            postcode = row[GeoNamesReader.COL_POSTCODE]
            place = row[GeoNamesReader.COL_PLACE].decode('utf-8')
            lat = row[GeoNamesReader.COL_LAT]
            lng = row[GeoNamesReader.COL_LNG]
            data.append((place, postcode, lat, lng))
        return data


class Mapper(object):
    def __init__(self, items1, items2):
        self.items1 = sorted(items1)
        self.items2 = sorted(items2)

    def merge(self):
        result = []
        for item1 in self.items1:
            place1, kshcode = item1
            for item2 in self.items2:
                place2, postcode, lat, lng = item2
                if place1 == place2:
                    result.append((place1, kshcode, postcode, lat, lng))
        return result


class CSVWriter(object):
    def __init__(self, filename):
        self.out = open(filename, 'w')

    def write(self, data):
        data = map(lambda lst: [elem.encode('utf-8') for elem in lst], data)
        writer = csv.writer(self.out)
        writer.writerows(data)


if __name__ == '__main__':
    ksh = KSHReader('Helysegnevkonyv_adattar_2010.xls')
    geonames = GeoNamesReader('HU.txt')
    mapper = Mapper(ksh.items(), geonames.items())
    data = mapper.merge()
    writer = CSVWriter('place_kshcode_postcode_points.csv')
    writer.write(data)
