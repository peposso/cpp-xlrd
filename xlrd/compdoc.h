#pragma once
// -*- coding: cp1252 -*-

////
// Implements the minimal functionality required
// to extract a "Workbook" or "Book" stream (as one big string);
// from an OLE2 Compound Document file.
// <p>Copyright � 2005-2012 Stephen John Machin, Lingfo Pty Ltd</p>
// <p>This module is part of the xlrd package, which is released under a BSD-style licence.</p>
////

// No part of the content of this file was derived from the works of David Giffin.

// 2008-11-04 SJM Avoid assertion error when -1 used instead of -2 for first_SID of empty SCSS [Frank Hoffsuemmer]
// 2007-09-08 SJM Warning message if sector sizes are extremely large.
// 2007-05-07 SJM Meaningful exception instead of IndexError if a SAT (sector allocation table) is corrupted.
// 2007-04-22 SJM Missing "<" in a struct.unpack call => can"t open files on bigendian platforms.

#include <string>
#include <exception>

#include "./utils.h"

namespace xlrd {
namespace compdoc {

const int DEBUG = 0;

using std::vector;
using u8 = uint8_t;

USING_FUNC(utils, pprint);
USING_FUNC(utils, slice);
USING_FUNC(utils, as_uint8);
USING_FUNC(utils, as_uint16);
USING_FUNC(utils, as_int32);
USING_FUNC(utils::str, format);
USING_FUNC(utils::str, unicode);

/*
from __future__ import pprint_function
import sys
from struct import unpack
from .timemachine import *
import array
*/

////
// Magic cookie that should appear in the first 8 bytes of the file.
const std::string SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

const int EOCSID = -2;
const int FREESID = -1;
const int SATSID = -3;
const int MSATSID = -4;
const int EVILSID = -5;

class CompDocError: public std::runtime_error
{
public:
    CompDocError(std::string msg) :std::runtime_error(msg) {}
    template<class...A>
    CompDocError(A...a) :std::runtime_error(format(a...)) {}
};


class DirNode
{
public:
    int DID;
    int etype;
    int colour;
    int left_DID;
    int right_DID;
    int root_DID;
    int first_SID;
    int tot_size;
    std::vector<int> children;
    int parent;
    std::vector<uint32_t> tsinfo;
    std::string name;

    DirNode(int DID, std::vector<uint8_t> dent, int DEBUG=0)
    {
        // dent is the 128-byte directory entry
        this->DID = DID;
        //(cbufsize, this->etype, this->colour, this->left_DID, this->right_DID,
        //this->root_DID) = 
        //    unpack("<HBBiii", dent[64:80]);
        utils::unpack bin64(dent, 64, 80);
        int cbufsize = bin64.as<uint16_t>();
        this->etype = bin64.as<uint8_t>();
        this->colour = bin64.as<uint8_t>();
        this->left_DID = bin64.as<int32_t>();
        this->right_DID = bin64.as<int32_t>();
        this->root_DID = bin64.as<int32_t>();
        //(this->first_SID, this->tot_size) = 
        //    unpack("<ii", dent[116:124]);
        utils::unpack bin116(dent, 116, 124);
        this->first_SID = bin116.as<int32_t>();
        this->tot_size = bin116.as<int32_t>();
        if (cbufsize == 0) {
            this->name = "";
        } else {
            this->name = unicode(slice(dent, 0, cbufsize-2), "utf_16_le");
        }
        // omit the trailing U+0000
        this->children = {}; // filled in later
        this->parent = -1; // indicates orphan; fixed up later
        //this->tsinfo = unpack("<IIII", dent[100:116]);
        utils::unpack bin100(dent, 100, 116);
        this->tsinfo = {bin100.as<uint32_t>(), bin100.as<uint32_t>(),
                        bin100.as<uint32_t>(), bin100.as<uint32_t>()};
        if (DEBUG) {
            this->dump(DEBUG);
        }
    }

    inline
    void dump(int DEBUG=1) {
        pprint(
            "DID=%d name=%s etype=%d DIDs(left=%d right=%d root=%d parent=%d kids=%s) first_SID=%d tot_size=%d\n",
            this->DID, this->name, this->etype, this->left_DID,
            this->right_DID, this->root_DID, this->parent, this->children, this->first_SID, this->tot_size
            );
        if (DEBUG == 2) {
            // cre_lo, cre_hi, mod_lo, mod_hi = tsinfo
            pprint("timestamp info %s", this->tsinfo);
        }
    }
};

inline
void _build_family_tree(std::vector<DirNode>& dirlist, int parent_DID, int child_DID) {
    if (child_DID < 0) return;
    _build_family_tree(dirlist, parent_DID, dirlist[child_DID].left_DID);
    dirlist[parent_DID].children.push_back(child_DID);
    dirlist[child_DID].parent = parent_DID;
    _build_family_tree(dirlist, parent_DID, dirlist[child_DID].right_DID);
    if (dirlist[child_DID].etype == 1) { // storage
        _build_family_tree(dirlist, child_DID, dirlist[child_DID].root_DID);
    }
}

////
// Compound document handler.
// @param mem The raw contents of the file, as a string, or as an mmap.mmap() object. The
// only operation it needs to support is slicing.

class CompDoc{
public:
    const std::vector<uint8_t>& mem;
    int sec_size;
    int short_sec_size;
    int dir_first_sec_sid;
    int min_size_std_stream;
    int mem_data_secs;
    int mem_data_len;
    vector<u8> seen;
    MAP<int, int> SAT;

    CompDoc(const std::vector<uint8_t>& mem)
    : mem(mem)
    {
        if (!utils::equals(slice(mem, 0, 8), SIGNATURE)) {
            throw CompDocError("Not an OLE2 compound document");
        }
        if (!utils::equals(slice(mem, 28, 30), "\xFE\xFF")) {
            throw CompDocError(format(
                "Expected \"little-endian\" marker, found %s",
                 slice(mem, 28, 30)));
        }
        //revision, version = unpack("<HH", mem[24:28]);
        int revision = as_uint16(mem, 24);
        int version = as_uint16(mem, 26);
        if (DEBUG) {
            pprint("\nCompDoc format: version=0x%04x revision=0x%04x",
                   version, revision);
        }
        // this->mem = mem;
        int ssz = as_uint16(mem, 30);
        int sssz = as_uint16(mem, 32);
        if (ssz > 20) { // allows for 2**20 bytes i.e. 1MB
            pprint("WARNING: sector size (2**%d) is preposterous; assuming 512 and continuing ...",
                   ssz);
            ssz = 9;
        }
        if (sssz > ssz) {
            pprint("WARNING: short stream sector size (2**%d) is preposterous; assuming 64 and continuing ...", sssz);
            sssz = 6;
        }
        int sec_size = 1 << ssz;
        this->sec_size = sec_size;
        this->short_sec_size = 1 << sssz;
        if (this->sec_size != 512 or this->short_sec_size != 64) {
            pprint("@@@@ sec_size=%d short_sec_size=%d",
                   this->sec_size, this->short_sec_size);
        }
        //(
        //    SAT_tot_secs, this->dir_first_sec_sid, _unused,
        //    this->min_size_std_stream,
        //    SSAT_first_sec_sid, SSAT_tot_secs,
        //    MSATX_first_sec_sid, MSATX_tot_secs,
        //// ) = unpack("<ii4xiiiii", mem[44:76]);
        //) = unpack("<iiiiiiii", mem[44:76]);
        int SAT_tot_secs          = as_int32(mem, 44);
        this->dir_first_sec_sid   = as_int32(mem, 48);
        int _unused               = as_int32(mem, 52);
        this->min_size_std_stream = as_int32(mem, 56);
        int SSAT_first_sec_sid    = as_int32(mem, 60);
        int SSAT_tot_secs         = as_int32(mem, 64);
        int MSATX_first_sec_sid   = as_int32(mem, 68);
        int MSATX_tot_secs        = as_int32(mem, 72);
        
        int mem_data_len = mem.size() - 512;
        int mem_data_secs = mem_data_len / sec_size;
        int left_over     = mem_data_len % sec_size;
        if (left_over) {
            //////// throw CompDocError("Not a whole number of sectors");
            mem_data_secs += 1;
            pprint("WARNING *** file size (%d) not 512 + multiple of sector size (%d)",
                   mem.size(), sec_size);
        }
        this->mem_data_secs = mem_data_secs; // use for checking later
        this->mem_data_len = mem_data_len;
        // seen = this->seen = array.array("B", [0]) * mem_data_secs;
        this->seen = vector<u8>(mem_data_secs);
        auto& seen = this->seen;

        if (DEBUG) {
            pprint("sec sizes", ssz, sssz, sec_size, this->short_sec_size);
            pprint("mem data: %d bytes == %d sectors", mem_data_len, mem_data_secs);
            pprint("SAT_tot_secs=%d, dir_first_sec_sid=%d, min_size_std_stream=%d",
                   SAT_tot_secs, this->dir_first_sec_sid, this->min_size_std_stream);
            pprint("SSAT_first_sec_sid=%d, SSAT_tot_secs=%d", SSAT_first_sec_sid, SSAT_tot_secs);
            pprint("MSATX_first_sec_sid=%d, MSATX_tot_secs=%d", MSATX_first_sec_sid, MSATX_tot_secs);
        }
        int nent = sec_size; // 4 // number of SID entries in a sector
        //fmt = "<%di" % nent
        int trunc_warned = 0;
        //
        // === build the MSAT ===
        //
        // MSAT = list(unpack("<109i", mem[76:512]));
        std::vector<int> MSAT;
        for (int i=0; i < 109; i ++) {
            MSAT.push_back(as_int32(mem, 76 + i*4));
        }
        int SAT_sectors_reqd = (mem_data_secs + nent - 1); // nent
        int expected_MSATX_sectors = std::max(0, (SAT_sectors_reqd - 109 + nent - 2)); // (nent - 1));
        int actual_MSATX_sectors = 0;
        if (MSATX_tot_secs == 0 and (MSATX_first_sec_sid == EOCSID ||
                                     MSATX_first_sec_sid == FREESID ||
                                     MSATX_first_sec_sid ==  0)) {
            // Strictly, if there is no MSAT extension, then MSATX_first_sec_sid
            // should be set to EOCSID ... FREESID and 0 have been met in the wild.
            //pass // Presuming no extension
        } else {
            int sid = MSATX_first_sec_sid;
            while (sid != EOCSID and sid != FREESID and sid != MSATSID) {
                // Above should be only EOCSID according to MS & OOo docs
                // but Excel doesn"t complain about FREESID. Zero is a valid
                // sector number, not a sentinel.
                if (DEBUG > 1) {
                    pprint("MSATX: sid=%d (0x%08X)", sid, sid);
                }
                if (sid >= mem_data_secs) {
                    std::string msg = format("MSAT extension: accessing sector %d but only %d in file", sid, mem_data_secs);
                    if (DEBUG > 1) {
                        pprint(msg);
                        break;
                    }
                    throw CompDocError(msg);
                } else if (sid < 0) {
                    throw CompDocError("MSAT extension: invalid sector id: %d", sid);
                }
                if (seen[sid]) {
                    throw CompDocError("MSAT corruption: seen[%d] == %d", sid, seen[sid]);
                }
                seen[sid] = 1;
                actual_MSATX_sectors += 1;
                if (DEBUG and actual_MSATX_sectors > expected_MSATX_sectors) {
                    pprint("[1]===>>>", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors);
                }
                int offset = 512 + sec_size * sid;
                // MSAT.extend(unpack(fmt, mem[offset:offset+sec_size]));
                for (int i=0; i < nent; ++i) {
                    MSAT.push_back(as_int32(mem, offset+i*4));
                }
                sid = MSAT.back(); // last sector id is sid of next sector in the chain
                MSAT.pop_back();
            }
        }
        if (DEBUG and actual_MSATX_sectors != expected_MSATX_sectors) {
            pprint("[2]===>>>", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors);
        }
        if (DEBUG) {
            pprint("MSAT: len = %lu", MSAT.size());
            dump_list(MSAT, 10);
        }
        //
        // === build the SAT ===
        //
        this->SAT = {};
        int actual_SAT_sectors = 0;
        int dump_again = 0;
        for (int msidx = 0, len = MSAT.size(); i < len; ++i) {
            int msid = MSAT[msidx];
            if (msid == FREESID or msid == EOCSID) {
                // Specification: the MSAT array may be padded with trailing FREESID entries.
                // Toleration: a FREESID or EOCSID entry anywhere in the MSAT array will be ignored.
                continue;
            }
            if (msid >= mem_data_secs) {
                if (not trunc_warned) {
                    pprint("WARNING *** File is truncated, or OLE2 MSAT is corrupt!!");
                    pprint("INFO: Trying to access sector %d but only %d available",
                           msid, mem_data_secs);
                    trunc_warned = 1;
                }
                MSAT[msidx] = EVILSID;
                dump_again = 1;
                continue;
            } else if (msid < -2) {
                throw CompDocError("MSAT: invalid sector id: %d", msid);
            }
            if (seen[msid]) {
                throw CompDocError("MSAT extension corruption: seen[%d] == %d", msid, seen[msid]);
            }
            seen[msid] = 2;
            actual_SAT_sectors += 1;
            if (DEBUG and actual_SAT_sectors > SAT_sectors_reqd) {
                pprint("[3]===>>>", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors, actual_SAT_sectors, msid);
            }
            offset = 512 + sec_size * msid;
            this->SAT.extend(unpack(fmt, mem[offset:offset+sec_size]));
        }

        if (DEBUG) {
            pprint("SAT: len = %lu", this->SAT.size());
            dump_list(this->SAT, 10);
            // pprint >> logfile, "SAT ",
            // for i, s in enumerate(this->SAT):
                // pprint >> logfile, "entry: %4d offset: %6d, next entry: %4d", i, 512 + sec_size * i, s);
                // pprint >> logfile, "%d:%d ", i, s),
        }
        if (DEBUG and dump_again) {
            pprint("MSAT: len =", MSAT.size());
            dump_list(MSAT, 10, logfile);
            for (int satx=mem_data_secs; i<this->SAT.size(), ++satx){
                this->SAT[satx] = EVILSID;
            }
            pprint("SAT: len = %lu", this->SAT.size());
            dump_list(this->SAT, 10);
        }
        //
        // === build the directory ===
        //
        dbytes = this->_get_stream(
            this->mem, 512, this->SAT, this->sec_size, this->dir_first_sec_sid,
            name="directory", seen_id=3);
        dirlist = [];
        did = -1;
        for pos in xrange(0, dbytes.size(), 128):
            did += 1;
            dirlist.append(DirNode(did, dbytes[pos:pos+128], 0, logfile));
        this->dirlist = dirlist;
        _build_family_tree(dirlist, 0, dirlist[0].root_DID); // and stand well back ...
        // if (DEBUG) {
        //     for (const auto& d: dirlist) {
        //         d.dump(DEBUG);
        //     }
        // }
        //
        // === get the SSCS ===
        //
        sscs_dir = this->dirlist[0];
        ASSERT(sscs_dir.etype == 5); // root entry
        if (sscs_dir.first_SID < 0 or sscs_dir.tot_size == 0) {
            // Problem reported by Frank Hoffsuemmer: some software was
            // writing -1 instead of -2 (EOCSID) for the first_SID
            // when the SCCS was empty. Not having EOCSID caused assertion
            // failure in _get_stream.
            // Solution: avoid calling _get_stream in any case when the
            // SCSS appears to be empty.
            this->SSCS = ""
        } else {
            this->SSCS = this->_get_stream(
                this->mem, 512, this->SAT, sec_size, sscs_dir.first_SID,
                sscs_dir.tot_size, name="SSCS", seen_id=4);
        }
        // if (DEBUG) { pprint >> logfile, "SSCS", repr(this->SSCS);
        //
        // === build the SSAT ===
        //
        this->SSAT.clear();
        if (SSAT_tot_secs > 0 and sscs_dir.tot_size == 0) {
            pprint("WARNING *** OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero");
        }
        if (sscs_dir.tot_size > 0) {
            int sid = SSAT_first_sec_sid;
            nsecs = SSAT_tot_secs;
            while (sid >= 0 and nsecs > 0) {
                if (seen[sid]) {
                    throw CompDocError("SSAT corruption: seen[%d] == %d", sid, seen[sid]);
                }
                seen[sid] = 5;
                nsecs -= 1;
                start_pos = 512 + sid * sec_size;
                news = list(unpack(fmt, mem[start_pos:start_pos+sec_size]));
                this->SSAT.extend(news);
                sid = this->SAT[sid];
            }
            if (DEBUG) { pprint("SSAT last sid %d; remaining sectors %d", sid, nsecs); }
            assert nsecs == 0 and sid == EOCSID
        if (DEBUG) {
            pprint("SSAT");
            dump_list(this->SSAT, 10, logfile);
        }
        if (DEBUG) {
            pprint("seen");
            dump_list(seen, 20, logfile);
        }
    }

    inline
    void _get_stream(vector<uint8_t> mem, int base, sat, sec_size, start_sid, size=None, name="", seen_id=None)
    {
        // pprint >> this->logfile, "_get_stream", base, sec_size, start_sid, size
        sectors = []
        s = start_sid
        if (size is None) {
            // nothing to check against
            while s >= 0:
                if (seen_id is not None) {
                    if (this->seen[s]) {
                        throw CompDocError("%s corruption: seen[%d] == %d", name, s, this->seen[s]));
                    this->seen[s] = seen_id
                start_pos = base + s * sec_size
                sectors.append(mem[start_pos:start_pos+sec_size]);
                try:
                    s = sat[s]
                except IndexError:
                    throw CompDocError(
                        "OLE2 stream %r: sector allocation table invalid entry (%d)" %
                        (name, s);
                        );
            assert s == EOCSID
        } else {
            todo = size;
            while (s >= 0) {
                if (seen_id is not None) {
                    if (this->seen[s]) {
                        throw CompDocError("%s corruption: seen[%d] == %d", name, s, this->seen[s]);
                    }
                    this->seen[s] = seen_id;
                }
                start_pos = base + s * sec_size;
                grab = sec_size;
                if (grab > todo) {
                    grab = todo;
                }
                todo -= grab;
                sectors.append(mem[start_pos:start_pos+grab]);
                try:
                    s = sat[s]
                except IndexError:
                    throw CompDocError(
                        "OLE2 stream %r: sector allocation table invalid entry (%d)",
                        name, s);
                }
            }
            assert s == EOCSID
            if (todo != 0) {
                pprint(this->logfile,
                    "WARNING *** OLE2 stream %r: expected size %d, actual size %d\n",
                    name, size, size - todo);
            }
        }
        return b"".join(sectors);
    }

    inline
    void _dir_search(path, storage_DID=0) {
        // Return matching DirNode instance, or None
        head = path[0]
        tail = path[1:]
        dl = this->dirlist
        for child in dl[storage_DID].children:
            if (dl[child].name.lower() == head.lower()) {
                et = dl[child].etype;
                if (et == 2) {
                    return dl[child];
                }
                if (et == 1) {
                    if (not tail) {
                        throw CompDocError("Requested component is a \"storage\"");
                    }
                    return this->_dir_search(tail, child);
                }
                dl[child].dump(1);
                throw CompDocError("Requested stream is not a \"user stream\"");
            }
        }
        return None
    }

    ////
    // Interrogate the compound document"s directory; return the stream as a string if found, otherwise
    // return None.
    // @param qname Name of the desired stream e.g. u"Workbook". Should be in Unicode or convertible thereto.

    inline
    void get_named_stream(this-> qname) {
        d = this->_dir_search(qname.split("/"));
        if (d is None) {
            return None;
        }
        if (d.tot_size >= this->min_size_std_stream) {
            return this->_get_stream(
                this->mem, 512, this->SAT, this->sec_size, d.first_SID,
                d.tot_size, name=qname, seen_id=d.DID+6);
        } else {
            return this->_get_stream(
                this->SSCS, 0, this->SSAT, this->short_sec_size, d.first_SID,
                d.tot_size, name=qname + " (from SSCS)", seen_id=None);
        }
    }

    ////
    // Interrogate the compound document"s directory.
    // If the named stream is not found, (None, 0, 0) will be returned.
    // If the named stream is found and is contiguous within the original byte sequence ("mem");
    // used when the document was opened,
    // then (mem, offset_to_start_of_stream, length_of_stream) is returned.
    // Otherwise a new string is built from the fragments and (new_string, 0, length_of_stream) is returned.
    // @param qname Name of the desired stream e.g. u"Workbook". Should be in Unicode or convertible thereto.

    void locate_named_stream(qname) {
        d = this->_dir_search(qname.split("/"));
        if (d is None) {
            return (None, 0, 0);
        if (d.tot_size > this->mem_data_len) {
            throw CompDocError("%s stream length (%d bytes) > file data size (%d bytes)"
               , qname, d.tot_size, this->mem_data_len));
        if (d.tot_size >= this->min_size_std_stream) {
            result = this->_locate_stream(
                this->mem, 512, this->SAT, this->sec_size, d.first_SID,
                d.tot_size, qname, d.DID+6);
            if (this->DEBUG) {
                pprint("\nseen", file=this->logfile);
                dump_list(this->seen, 20, this->logfile);
            return result
        } else {
            return (
                this->_get_stream(
                    this->SSCS, 0, this->SSAT, this->short_sec_size, d.first_SID,
                    d.tot_size, qname + " (from SSCS)", None),
                0,
                d.tot_size
                );
        }
    }

    void _locate_stream(std::vector<uint8_t> mem, int base,
                        sat, sec_size,
                        int start_sid, int expected_stream_size,
                        std::string qname, int seen_id)
    {
        // pprint >> this->logfile, "_locate_stream", base, sec_size, start_sid, expected_stream_size
        int s = start_sid;
        if (s < 0) {
            throw CompDocError(
                "_locate_stream: start_sid (%d) is -ve", start_sid);
        }
        int p = -99; // dummy previous SID
        int start_pos = -9999;
        int end_pos = -8888;
        vector<array<int, 2>> slices;
        int tot_found = 0;
        int found_limit = (expected_stream_size + sec_size - 1); // sec_size
        while (s >= 0) {
            if (this->seen[s]) {
                pprint("_locate_stream(%s): seen", qname);
                dump_list(this->seen, 20);
                throw CompDocError("%s corruption: seen[%d] == %d", qname, s, this->seen[s]);
            }
            this->seen[s] = seen_id;
            tot_found += 1;
            if (tot_found > found_limit) {
                throw CompDocError(
                    "%s: size exceeds expected %d bytes; corrupt?",
                    qname, found_limit * sec_size);
                    // Note: expected size rounded up to higher sector
            }
            if (s == p+1) {
                // contiguous sectors
                end_pos += sec_size;
            } else {
                // start new slice
                if (p >= 0) {
                    // not first time
                    slices.push_back({{start_pos, end_pos}});
                }
                start_pos = base + s * sec_size;
                end_pos = start_pos + sec_size;
            }
            p = s;
            s = sat[s];
        }
        ASSERT(s == EOCSID);
        ASSERT(tot_found == found_limit);
        // pprint >> this->logfile, "_locate_stream(%s): seen" % qname; dump_list(this->seen, 20, this->logfile);
        if (slices.empty()) {
            // The stream is contiguous ... just what we like!
            return (mem, start_pos, expected_stream_size);
        }
        slices.append((start_pos, end_pos));
        // pprint >> this->logfile, "+++>>> %d fragments" % slices.size();
        return (b"".join([mem[start_pos:end_pos] for start_pos, end_pos in slices]), 0, expected_stream_size);
    }
};

// ==========================================================================================
/*
def x_dump_line(alist, stride, f, dpos, equal=0):
    pprint("%5d%s", dpos, " ="[equal]), end=" ", file=f);
    for value in alist[dpos:dpos + stride]:
        pprint(str(value), end=" ", file=f);
    pprint(file=f);

def dump_list(alist, stride, f=sys.stdout):
    def _dump_line(dpos, equal=0):
        pprint("%5d%s", dpos, " ="[equal]), end=" ", file=f);
        for value in alist[dpos:dpos + stride]:
            pprint(str(value), end=" ", file=f);
        pprint(file=f);
    pos = None
    oldpos = None
    for pos in xrange(0, alist.size(), stride):
        if (oldpos is None) {
            _dump_line(pos);
            oldpos = pos
        } else if (alist[pos:pos+stride] != alist[oldpos:oldpos+stride]) {
            if (pos - oldpos > stride) {
                _dump_line(pos - stride, equal=1);
            _dump_line(pos);
            oldpos = pos
    if (oldpos is not None and pos is not None and pos != oldpos) {
        _dump_line(pos, equal=1);
*/

}
}

