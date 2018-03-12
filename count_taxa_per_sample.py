#!/usr/bin/python

import sys, os, argparse
import glob
import tempfile
import subprocess
import itertools
from collections import Counter
import xlsxwriter

def count_higher_taxon_per_sample(taxon):
    with open(sys.argv[1]) as header:
        l = header.readline().strip()
        samples = l.split("\t")[1:]
    count = 1
    count_persample = {}
    for x in samples:
        taxon_family = {}
        with open(sys.argv[1]) as otutable:
            for row in otutable:
                if row[0] != "#":
                    try:
                        row = row.split("\t")
                        taxonlist = map(str.strip, row[-1].split("/"))
                        #print taxonlist
                        if len(taxonlist) > 5:
                            if taxonlist[taxon] in taxon_family:
                                taxon_family[taxonlist[taxon]] += int(row[count])
                            else:
                                taxon_family[taxonlist[taxon]] = int(row[count])
                        else:
                            if "no identification" in taxon_family:
                                taxon_family[taxonlist[taxon]] += int(row[count])
                            else:
                                taxon_family["no identification"] = int(row[count])
                            
                    except:
                        #print row
                        pass
        count += 1
        count_persample[str(x)] = taxon_family
    #print taxon_family
    return count_persample, taxon_family

def write_to_excel(data, names, worksheets, taxon):
    #header
    samplenames = ""
    col = 1
    for samplename in data:
        worksheets["worksheet_"+taxon].write(0, col, samplename)
        col += 1
  
    row = 1
    for name in names:
        col = 0
        line = name
        worksheets["worksheet_"+taxon].write(row, col, name)
        #print name
        for count in data:
            col += 1
            worksheets["worksheet_"+taxon].write(row, col, data[count][name])    
        row += 1
    
def main():
    worksheets = {}
    workbook = xlsxwriter.Workbook(sys.argv[2])
    #worksheets["worksheet_otu_table"] = workbook.add_worksheet("otu table")
    worksheets["worksheet_species"] = workbook.add_worksheet("species")
    worksheets["worksheet_genus"] = workbook.add_worksheet("genus")
    worksheets["worksheet_family"] = workbook.add_worksheet("family")
    worksheets["worksheet_order"] = workbook.add_worksheet("order")
    worksheets["worksheet_class"] = workbook.add_worksheet("class")
    worksheets["worksheet_phylum"] = workbook.add_worksheet("phylum")
    worksheets["worksheet_kingdom"] = workbook.add_worksheet("kingdom")
    #worksheets[worksheet_otu_table].write('A1', 'Hello world')
    
    taxon_index = {'kingdom':0, 'phylum':1, 'class':2, 'order':3, 'family':4, 'genus':5, 'species':6}
    """
    data, names = count_higher_taxon_per_sample(taxon_index['kingdom'])
    write_to_excel(data, names, worksheets, "kingdom")
    """
    for taxon in taxon_index:
        data, names = count_higher_taxon_per_sample(taxon_index[taxon])
        write_to_excel(data, names, worksheets, taxon)
    


    print "done"
    


if __name__ == '__main__':
    main()
