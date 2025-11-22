/**
 * OLE/MIME Email File Parser
 * * processes .msg (OLE Compound Document) and .eml (MIME) file formats.
 * Handles binary property extraction, text decoding, and recipient reconciliation.
 */

'use strict';

// MAPI Property Tags
const PROP_TYPE_INTEGER32 = 0x0003;
const PROP_TYPE_BOOLEAN = 0x000B;
const PROP_TYPE_STRING = 0x001E;
const PROP_TYPE_STRING8 = 0x001F;
const PROP_TYPE_TIME = 0x0040;
const PROP_TYPE_BINARY = 0x0102;

const PROP_ID_SUBJECT = 0x0037;
const PROP_ID_BODY = 0x1000;
const PROP_ID_HTML_BODY = 0x1013;
const PROP_ID_DISPLAY_TO = 0x0E04;
const PROP_ID_DISPLAY_CC = 0x0E03;
const PROP_ID_DISPLAY_BCC = 0x0E02;

const PROP_ID_RECIPIENT_TYPE = 0x0C15;
const PROP_ID_RECIPIENT_DISPLAY_NAME = 0x3001;
const PROP_ID_RECIPIENT_EMAIL_ADDRESS = 0x3003;
const PROP_ID_RECIPIENT_SMTP_ADDRESS = 0x39FE;

const RECIPIENT_TYPE_TO = 1;
const RECIPIENT_TYPE_CC = 2;
const RECIPIENT_TYPE_BCC = 3;

// Module-Level Decoders
let _textDecoderUtf16 = null;
let _textDecoderWin1252 = null;
let _domParser = null;

function getTextDecoder(encoding) {
    if (encoding === 'utf-16le') {
        if (!_textDecoderUtf16) _textDecoderUtf16 = new TextDecoder('utf-16le', { fatal: false });
        return _textDecoderUtf16;
    }
    if (encoding === 'windows-1252') {
        if (!_textDecoderWin1252) _textDecoderWin1252 = new TextDecoder('windows-1252', { fatal: false });
        return _textDecoderWin1252;
    }
    return new TextDecoder(encoding, { fatal: false });
}

function getDOMParser() {
    if (!_domParser && typeof DOMParser !== 'undefined') {
        _domParser = new DOMParser();
    }
    return _domParser;
}

/* =============================================================================
   TEXT PROCESSING UTILS
   ============================================================================= */

/**
 * Decodes Quoted-Printable strings to raw text.
 */
function _decodeQuotedPrintable(str, charset = 'utf-8') {
    if (!str) return '';
    
    // Join soft lines and convert hex to chars
    let decoded = str
        .replace(/=(\r\n|\n)/g, '') 
        .replace(/=([0-9A-F]{2})/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
    
    try {
        let bytes = new Uint8Array(decoded.length);
        for (let i = 0; i < decoded.length; i++) bytes[i] = decoded.charCodeAt(i);
        
        let encoding = charset.toLowerCase();
        if (encoding === 'us-ascii') encoding = 'utf-8';
        
        return new TextDecoder(encoding, { fatal: false }).decode(bytes);
    } catch (e) { 
        return decoded; 
    }
}

/**
 * Removes HTML tags and CSS artifacts, normalizing to plain text.
 */
function _stripHtml(html) {
    if (!html || typeof html !== 'string') return ''; 
    
    // Aggressive Regex Removal for scripts/styles
    let text = html
        .replace(/<head[\s\S]*?<\/head>/gi, '')
        .replace(/<style[\s\S]*?<\/style>/gi, '')
        .replace(/<script[\s\S]*?<\/script>/gi, '')
        .replace(/<!--[\s\S]*?-->/g, '')
        .replace(/<\?xml[^>]*\?>/gi, '');

    // Outlook-specific CSS artifact removal
    text = text.replace(/[.#a-z0-9_]+\s*\{[^}]+\}/gi, '');

    // Block tag spacing
    text = text.replace(/\s*<(br|p|div|tr|li|h1|h2|h3|h4|h5|h6)[^>]*>\s*/gi, '\n');
    
    let parser = getDOMParser();
    if (parser) {
        try {
            let doc = parser.parseFromString(text, 'text/html');
            const junk = doc.querySelectorAll('style, script, link, meta, title');
            junk.forEach(el => el.remove());
            text = doc.body ? doc.body.textContent : (doc.documentElement.textContent || '');
        } catch (e) {
            text = text.replace(/<[^>]+>/g, '');
        }
    } else {
        text = text.replace(/<[^>]+>/g, '');
    }
    
    return _normalizeText(text);
}

/**
 * Normalizes line endings and collapses excessive whitespace.
 */
function _normalizeText(text) {
    if (!text) return '';
    return text
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/\n{3,}/g, '\n\n') 
        .trim();
}

function dataViewToString(view, encoding) {
    if (encoding === 'utf-8') {
        try {
            if (typeof TextDecoder === 'undefined') throw new Error("TextDecoder missing");
            let decoded = new TextDecoder('utf-8', { fatal: true }).decode(view);
            const nullIdx = decoded.indexOf('\0');
            return nullIdx !== -1 ? decoded.substring(0, nullIdx) : decoded;
        } catch (e) { 
            return dataViewToString(view, 'ascii'); 
        }
    }
    
    if (encoding === 'utf16le') {
        try {
            if (typeof TextDecoder === 'undefined') throw new Error("TextDecoder missing");
            let decoded = getTextDecoder('utf-16le').decode(view);
            const nullIdx = decoded.indexOf('\0');
            return nullIdx !== -1 ? decoded.substring(0, nullIdx) : decoded;
        } catch (e) {
            let result = '';
            for (let i = 0; i < view.byteLength - 1; i += 2) {
                let charCode = view.getUint16(i, true);
                if (charCode === 0) break;
                result += String.fromCharCode(charCode);
            }
            return result;
        }
    }
    
    try {
        let decoded = getTextDecoder('windows-1252').decode(view);
        const nullIdx = decoded.indexOf('\0');
        return nullIdx !== -1 ? decoded.substring(0, nullIdx) : decoded;
    } catch(e) {
        let result = '';
        for (let i = 0; i < view.byteLength; i++) {
            let charCode = view.getUint8(i);
            if (charCode === 0) break;
            result += String.fromCharCode(charCode);
        }
        return result;
    }
}

function filetimeToDate(low, high) {
    if (typeof BigInt === 'undefined') return null;
    try {
        const FILETIME_EPOCH_DIFF = 116444736000000000n;
        let filetime = (BigInt(high) << 32n) | BigInt(low);
        return new Date(Number((filetime - FILETIME_EPOCH_DIFF) / 10000n));
    } catch (e) { return null; }
}

/* =============================================================================
   PARSER LOGIC
   ============================================================================= */

function _parsePropTag(entryName) {
    let propTagStr = "00000000";
    if (entryName.length >= 20) propTagStr = entryName.substring(entryName.length - 8);
    else {
        let parts = entryName.split('_');
        if (parts.length >= 3) propTagStr = parts[2];
        else return null;
    }
    try {
        return { id: parseInt(propTagStr.substring(0, 4), 16), type: parseInt(propTagStr.substring(4, 8), 16) };
    } catch (e) { return null; }
}

function _shouldStoreProperty(propId, newPropType, existingProp) {
    let isBodyProperty = (propId === PROP_ID_BODY || propId === PROP_ID_HTML_BODY);
    if (!isBodyProperty || !existingProp) return true;
    
    // Prefer String types over other types for body
    let existingIsText = (existingProp.type === PROP_TYPE_STRING || existingProp.type === PROP_TYPE_STRING8);
    let newIsText = (newPropType === PROP_TYPE_STRING || newPropType === PROP_TYPE_STRING8);
    
    if (existingIsText && !newIsText) return false;
    if (!existingIsText && newIsText) return true;
    return false;
}

function parseAddress(addr) {
    if (!addr) return { name: '', email: null };
    addr = addr.trim();
    let email = addr, name = addr;
    let match = addr.match(/^(.*)<([^>]+)>$/);
    if (match) { name = match[1].trim().replace(/^"|"$/g, ''); email = match[2].trim(); }
    let emailMatch = email.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
    if (emailMatch) {
        email = emailMatch[0];
        if (name === addr) name = (name === email) ? '' : name.replace(email, '').trim();
    } else email = null;
    return { name, email };
}

/**
 * Main Parser Class
 */
function MsgReaderParser(arrayBuffer) {
    if (!(arrayBuffer instanceof ArrayBuffer) && !(arrayBuffer instanceof Uint8Array)) {
        throw new Error("MsgReader: Input must be ArrayBuffer or Uint8Array.");
    }
    this.buffer = arrayBuffer instanceof ArrayBuffer ? arrayBuffer : new Uint8Array(arrayBuffer).buffer;
    this.dataView = new DataView(this.buffer);
    this.header = null; this.fat = null; this.miniFat = null;
    this.directoryEntries = []; this.properties = {}; this._mimeScanCache = null;
}

MsgReaderParser.prototype.parse = function() {
    this.readHeader(); this.readFAT(); this.readMiniFAT(); this.readDirectory(); this.extractProperties();
    return {
        getFieldValue: this.getFieldValue.bind(this),
        subject: this.getFieldValue('subject'),
        body: this.getFieldValue('body'),
        bodyHTML: this.getFieldValue('bodyHTML'),
        recipients: this.getFieldValue('recipients')
    };
};

MsgReaderParser.prototype.parseMime = function() {
    this._mimeScanCache = null;
    let rawText = '';
    try { rawText = new TextDecoder('utf-8', { fatal: false }).decode(this.dataView); }
    catch (e) { 
        try { rawText = new TextDecoder('latin1').decode(this.dataView); }
        catch (e2) { rawText = ''; }
    }
    
    let mimeData = this._scanBufferForMimeText(rawText);
    let recipients = [];
    
    let parseMimeAddresses = (addrString, type) => {
        if (!addrString) return;
        addrString.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).forEach(addr => {
            let parsed = parseAddress(addr);
            if (parsed.email) recipients.push({ name: parsed.name, email: parsed.email, recipientType: type });
        });
    };

    parseMimeAddresses(mimeData.to, RECIPIENT_TYPE_TO);
    parseMimeAddresses(mimeData.cc, RECIPIENT_TYPE_CC);
    let bccMatch = rawText.match(/^Bcc:\s*([^\r\n]+)/im);
    if (bccMatch) parseMimeAddresses(bccMatch[1].trim(), RECIPIENT_TYPE_BCC);

    this.properties[PROP_ID_SUBJECT] = { id: PROP_ID_SUBJECT, value: mimeData.subject };
    this.properties[PROP_ID_BODY] = { id: PROP_ID_BODY, value: mimeData.body };
    this.properties[PROP_ID_HTML_BODY] = { id: PROP_ID_HTML_BODY, value: null };
    this.properties['recipients'] = { id: 0, value: recipients };

    return {
        getFieldValue: this.getFieldValue.bind(this),
        subject: mimeData.subject,
        body: mimeData.body,
        bodyHTML: null, recipients: recipients
    };
};

MsgReaderParser.prototype.readHeader = function() {
    if (this.buffer.byteLength < 512) throw new Error('File too small to be a valid OLE file');
    if (this.dataView.getUint32(0, true) !== 0xE011CFD0) throw new Error('Invalid OLE file signature.');
    this.header = {
        sectorShift: this.dataView.getUint16(30, true),
        miniSectorShift: this.dataView.getUint16(32, true),
        fatSectors: this.dataView.getUint32(44, true),
        directoryFirstSector: this.dataView.getUint32(48, true),
        miniFatFirstSector: this.dataView.getUint32(60, true),
        miniFatTotalSectors: this.dataView.getUint32(64, true),
        difFirstSector: this.dataView.getUint32(68, true),
        difTotalSectors: this.dataView.getUint32(72, true)
    };
    this.header.sectorSize = Math.pow(2, this.header.sectorShift);
    this.header.miniSectorSize = Math.pow(2, this.header.miniSectorShift);
};

MsgReaderParser.prototype.readFAT = function() {
    let sectorSize = this.header.sectorSize, entriesPerSector = sectorSize / 4;
    this.fat = [];
    let fatSectorPositions = [];
    for (let i = 0; i < 109 && i < this.header.fatSectors; i++) {
        let s = this.dataView.getUint32(76 + i * 4, true);
        if (s !== 0xFFFFFFFE && s !== 0xFFFFFFFF) fatSectorPositions.push(s);
    }
    if (this.header.difTotalSectors > 0) {
        let difSector = this.header.difFirstSector;
        let sectorsRead = 0;
        while (difSector !== 0xFFFFFFFE && difSector !== 0xFFFFFFFF && sectorsRead < this.header.difTotalSectors) {
            let difOffset = 512 + difSector * sectorSize;
            for (let j = 0; j < entriesPerSector - 1; j++) {
                let s = this.dataView.getUint32(difOffset + j * 4, true);
                if (s !== 0xFFFFFFFE && s !== 0xFFFFFFFF) fatSectorPositions.push(s);
            }
            difSector = this.dataView.getUint32(difOffset + (entriesPerSector - 1) * 4, true);
            sectorsRead++;
        }
    }
    for (let i = 0; i < fatSectorPositions.length; i++) {
        let offset = 512 + fatSectorPositions[i] * sectorSize;
        for (let j = 0; j < entriesPerSector; j++) {
            if (offset + j * 4 + 4 <= this.buffer.byteLength) this.fat.push(this.dataView.getUint32(offset + j * 4, true));
        }
    }
};

MsgReaderParser.prototype.readMiniFAT = function() {
    if (this.header.miniFatFirstSector === 0xFFFFFFFE) { this.miniFat = []; return; }
    this.miniFat = [];
    let sector = this.header.miniFatFirstSector, sectorSize = this.header.sectorSize;
    let sectorsRead = 0; 
    while (sector !== 0xFFFFFFFE && sector !== 0xFFFFFFFF && sectorsRead < this.header.miniFatTotalSectors) {
        let offset = 512 + sector * sectorSize;
        for (let i = 0; i < sectorSize / 4; i++) {
            if (offset + i * 4 + 4 <= this.buffer.byteLength) this.miniFat.push(this.dataView.getUint32(offset + i * 4, true));
        }
        if (sector >= this.fat.length) break;
        sector = this.fat[sector];
        sectorsRead++;
    }
};

MsgReaderParser.prototype.readDirectory = function() {
    let sector = this.header.directoryFirstSector, sectorSize = this.header.sectorSize, entrySize = 128;
    let sectorsRead = 0;
    while (sector !== 0xFFFFFFFE && sector !== 0xFFFFFFFF) {
        let offset = 512 + sector * sectorSize;
        for (let i = 0; i < sectorSize / entrySize; i++) {
            let entryOffset = offset + i * entrySize;
            if (entryOffset + entrySize > this.buffer.byteLength) break;
            let entry = this.readDirectoryEntry(entryOffset);
            if (entry && entry.name) this.directoryEntries.push(entry);
        }
        if (sector >= this.fat.length) break;
        sector = this.fat[sector];
        sectorsRead++;
    }
    this.directoryEntries.forEach((de, idx) => de.id = idx);
};

MsgReaderParser.prototype.readDirectoryEntry = function(offset) {
    let nameLen = this.dataView.getUint16(offset + 64, true);
    if (nameLen === 0 || nameLen > 64) return null;
    let name = dataViewToString(new DataView(this.buffer, offset, Math.min(nameLen, 64)), 'utf16le');
    let type = this.dataView.getUint8(offset + 66);
    if (type !== 1 && type !== 2 && type !== 5) return null;
    return {
        name: name, type: type,
        startSector: this.dataView.getUint32(offset + 116, true),
        size: this.dataView.getUint32(offset + 120, true),
        leftSiblingId: this.dataView.getInt32(offset + 68, true),
        rightSiblingId: this.dataView.getInt32(offset + 72, true),
        childId: this.dataView.getInt32(offset + 76, true),
        id: -1
    };
};

MsgReaderParser.prototype._readSectorChain = function(startSector, sectorSize, fatArray, totalSize) {
    let data = new Uint8Array(totalSize);
    let dataOffset = 0;
    let sector = startSector;
    
    while (sector !== 0xFFFFFFFE && sector !== 0xFFFFFFFF && dataOffset < totalSize) {
        let offset = (fatArray === this.miniFat) 
            ? sector * sectorSize 
            : 512 + sector * sectorSize;
            
        let sourceData = (fatArray === this.miniFat)
            ? this._miniStreamData
            : this.dataView; 
            
        let copy = Math.min(sectorSize, totalSize - dataOffset);
        
        for (let i = 0; i < copy; i++) {
            data[dataOffset++] = (fatArray === this.miniFat) 
                ? sourceData[offset + i] 
                : sourceData.getUint8(offset + i);
        }
        
        if (sector >= fatArray.length) break;
        sector = fatArray[sector];
    }
    return data;
};

MsgReaderParser.prototype.readStream = function(entry) {
    if (!entry || entry.size === 0) return new Uint8Array(0);
    
    if (entry.size < 4096 && entry.type !== 5) {
        let root = this.directoryEntries.find(e => e.type === 5);
        if (!root) return new Uint8Array(0);
        
        if (!this._miniStreamData) {
             this._miniStreamData = this.readStream(root);
        }
        
        return this._readSectorChain(entry.startSector, this.header.miniSectorSize, this.miniFat, entry.size);
    } else {
        return this._readSectorChain(entry.startSector, this.header.sectorSize, this.fat, entry.size);
    }
};

MsgReaderParser.prototype._scanBufferForMimeText = function(rawText) {
    if (this._mimeScanCache) return this._mimeScanCache;
    
    if (!rawText) {
        try { rawText = new TextDecoder('utf-8', { fatal: false }).decode(this.dataView); }
        catch (e) { 
            try { rawText = new TextDecoder('latin1').decode(this.dataView); }
            catch (e2) { 
                return { subject: null, to: null, cc: null, body: null };
            }
        }
    }

    let result = { subject: null, to: null, cc: null, body: null };
    
    const findField = (name) => {
        const search = new RegExp(`\\b${name}:\\s*([^\\r\\n]+)`, 'i');
        const match = rawText.match(search);
        return match ? match[1].trim() : null;
    };
    
    result.subject = findField('Subject');
    result.to = findField('To');
    result.cc = findField('Cc');
    
    let headerEndIndex = rawText.indexOf('\r\n\r\n');
    if (headerEndIndex === -1) headerEndIndex = rawText.indexOf('\n\n');
    
    if (headerEndIndex !== -1) {
        let bodyText = rawText.substring(headerEndIndex + 4);
        let encoding = null;
        let charset = 'utf-8'; // Default

        if (/Content-Type:\s*multipart/i.test(rawText)) {
             const plainMatch = bodyText.match(/Content-Type:\s*text\/plain/i);
             if (plainMatch) {
                 const plainTypeIndex = plainMatch.index;
                 
                 let start = -1;
                 let startOffset = 0;
                 
                 const rnrn = bodyText.indexOf('\r\n\r\n', plainTypeIndex);
                 const nn = bodyText.indexOf('\n\n', plainTypeIndex);
                 
                 if (rnrn !== -1 && (nn === -1 || rnrn < nn)) {
                     start = rnrn;
                     startOffset = 4;
                 } else if (nn !== -1) {
                     start = nn;
                     startOffset = 2;
                 }
                 
                 const partHeaders = bodyText.substring(plainTypeIndex, start);
                 if (/Content-Transfer-Encoding:\s*quoted-printable/i.test(partHeaders)) {
                     encoding = 'quoted-printable';
                 }
                 
                 const charsetMatch = partHeaders.match(/charset=["']?([^"';\r\n]+)/i);
                 if (charsetMatch) charset = charsetMatch[1];

                 if (start !== -1) {
                     let end = -1;
                     const boundRN = bodyText.indexOf('\r\n--', start + startOffset);
                     const boundN = bodyText.indexOf('\n--', start + startOffset);
                     
                     if (boundRN !== -1 && (boundN === -1 || boundRN < boundN)) {
                         end = boundRN;
                     } else if (boundN !== -1) {
                         end = boundN;
                     }
                     
                     if (end === -1) end = bodyText.length;
                     
                     bodyText = bodyText.substring(start + startOffset, end).trim();
                 }
             }
        } else {
             if (rawText.match(/^Content-Transfer-Encoding:\s*quoted-printable/im)) {
                encoding = 'quoted-printable';
             }
             // Check global charset
             const charsetMatch = rawText.match(/charset=["']?([^"';\r\n]+)/i);
             if (charsetMatch) charset = charsetMatch[1];
        }
        
        result.body = (encoding === 'quoted-printable') ? _decodeQuotedPrintable(bodyText, charset) : bodyText;
    }
    
    this._mimeScanCache = result;
    return result;
};

MsgReaderParser.prototype.extractProperties = function() {
    let self = this, rawProps = {};
    this.directoryEntries.forEach(entry => {
        if (entry.name.indexOf('__substg1.0_') !== 0 || entry.name.indexOf('__recip_version1.0_') > -1) return;
        let propTag = _parsePropTag(entry.name);
        if (!propTag) return;
        if (!_shouldStoreProperty(propTag.id, propTag.type, rawProps[propTag.id])) return;
        rawProps[propTag.id] = { id: propTag.id, type: propTag.type, data: self.readStream(entry) };
    });

    let getVal = (id, type) => {
        let p = rawProps[id];
        return p ? self.convertPropertyValue(p.data, p.type, id) : null;
    };

    let bodyHtml = getVal(PROP_ID_HTML_BODY, PROP_TYPE_STRING);
    if (bodyHtml) this.properties[PROP_ID_HTML_BODY] = { id: PROP_ID_HTML_BODY, value: bodyHtml };
    
    if (this.properties[PROP_ID_HTML_BODY] && this.properties[PROP_ID_HTML_BODY].value instanceof Uint8Array) {
        let view = new DataView(this.properties[PROP_ID_HTML_BODY].value.buffer);
        let decoded = dataViewToString(view, 'utf-8');
        this.properties[PROP_ID_HTML_BODY].value = decoded;
        bodyHtml = decoded;
    }

    let body = getVal(PROP_ID_BODY, PROP_TYPE_STRING);
    
    if (body instanceof Uint8Array) {
         let view = new DataView(body.buffer);
         body = dataViewToString(view, 'utf-8');
         this.properties[PROP_ID_BODY] = { id: PROP_ID_BODY, value: body };
    }

    if (!body && bodyHtml) body = _stripHtml(bodyHtml);
    if (body) this.properties[PROP_ID_BODY] = { id: PROP_ID_BODY, value: body };

    Object.values(rawProps).forEach(p => {
        if (p.id !== PROP_ID_BODY && p.id !== PROP_ID_HTML_BODY) {
            this.properties[p.id] = { id: p.id, value: self.convertPropertyValue(p.data, p.type, p.id) };
        }
    });

    let mimeData = this._scanBufferForMimeText(null);
    if (!this.properties[PROP_ID_SUBJECT]) this.properties[PROP_ID_SUBJECT] = { value: mimeData.subject };
    if (!this.properties[PROP_ID_BODY]) this.properties[PROP_ID_BODY] = { value: mimeData.body };

    this.extractRecipients();
};

MsgReaderParser.prototype.convertPropertyValue = function(data, type, propId) {
    if (!data || data.length === 0) return null;
    let view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    
    const isBodyProp = (propId === PROP_ID_BODY || propId === PROP_ID_HTML_BODY);
    
    if (isBodyProp || type === PROP_TYPE_STRING || type === PROP_TYPE_STRING8) {
        let u16 = '', u8 = '';
        try { u16 = dataViewToString(view, 'utf16le'); } catch (e) {}
        try { u8 = dataViewToString(view, 'utf-8'); } catch (e) {}
        
        let isPrintable = (s) => {
            if (!s || s.length === 0) return false;
            let printableCount = s.replace(/[^\x20-\x7E\n\r\t\u00A0-\u00FF]/g, '').length;
            return (printableCount / s.length) > 0.7;
        };
        
        let u16IsBetter = isPrintable(u16);
        let u8IsBetter = isPrintable(u8);
        
        if (u8IsBetter && u16IsBetter && u8.length < u16.length && u8.length < 5) {
             u8IsBetter = false;
        }
        
        let useU16 = false;
        if (type === PROP_TYPE_STRING8) {
            useU16 = u16IsBetter && !u8IsBetter;
        } else {
            useU16 = u16IsBetter;
        }
        
        let text = useU16 ? u16 : u8;

        if (propId === PROP_ID_BODY) return _normalizeText(_stripHtml(text));
        
        return text;
    }
    
    if (type === PROP_TYPE_BINARY) return data;
    if (type === PROP_TYPE_INTEGER32) return view.byteLength >= 4 ? view.getUint32(0, true) : 0;
    if (type === PROP_TYPE_BOOLEAN) return view.byteLength > 0 ? view.getUint8(0) !== 0 : false;
    if (type === PROP_TYPE_TIME) return view.byteLength >= 8 ? filetimeToDate(view.getUint32(0, true), view.getUint32(4, true)) : null;
    return data;
};

function _extractAddresses(displayString) {
    let emails = [];
    if (displayString) {
        displayString.split(/[;,]/).forEach(addr => {
            let parsed = parseAddress(addr);
            if (parsed.email) emails.push(parsed.email.toLowerCase());
        });
    }
    return emails;
}

MsgReaderParser.prototype.extractRecipients = function() {
    let self = this;
    let recipients = [];
    
    let recipientStorages = this.directoryEntries.filter(entry => 
        entry.type === 1 && entry.name.indexOf('__recip_version1.0_') === 0
    );
    
    recipientStorages.forEach(storage => {
        let recipient = {
            recipientType: RECIPIENT_TYPE_TO,
            name: '',
            email: ''
        };
        
        let findChildren = (parentId) => {
            let parent = self.directoryEntries[parentId];
            if (!parent || parent.childId === -1) return [];
            let children = [];
            let stack = [parent.childId];
            let visited = new Set();
            
            while (stack.length > 0) {
                let id = stack.pop();
                if (id === -1 || visited.has(id)) continue;
                visited.add(id);
                let entry = self.directoryEntries[id];
                if (!entry) continue;
                children.push(entry);
                if (entry.leftSiblingId !== -1) stack.push(entry.leftSiblingId);
                if (entry.rightSiblingId !== -1) stack.push(entry.rightSiblingId);
            }
            return children;
        };
        
        let children = findChildren(storage.id);
        children.forEach(child => {
            let propTag = _parsePropTag(child.name);
            if (!propTag) return;
            
            let propData = self.readStream(child);
            let propValue = self.convertPropertyValue(propData, propTag.type, propTag.id);
            
            if (propTag.id === PROP_ID_RECIPIENT_TYPE) {
                recipient.recipientType = propValue || RECIPIENT_TYPE_TO;
            } else if (propTag.id === PROP_ID_RECIPIENT_DISPLAY_NAME) {
                recipient.name = propValue || '';
            } else if (propTag.id === PROP_ID_RECIPIENT_EMAIL_ADDRESS || propTag.id === PROP_ID_RECIPIENT_SMTP_ADDRESS) {
                if (propValue && propValue.indexOf('@') > -1) {
                    recipient.email = propValue;
                }
            }
        });
        
        if (recipient.email || recipient.name) {
            recipients.push(recipient);
        }
    });
    
    let displayTo = this.properties[PROP_ID_DISPLAY_TO] ? this.properties[PROP_ID_DISPLAY_TO].value : null;
    let displayCc = this.properties[PROP_ID_DISPLAY_CC] ? this.properties[PROP_ID_DISPLAY_CC].value : null;
    
    if (displayTo || displayCc) {
        let displayToEmails = _extractAddresses(displayTo);
        let displayCcEmails = _extractAddresses(displayCc);
        
        let toEmailCounts = {};
        let ccEmailCounts = {};
        
        displayToEmails.forEach(email => {
            toEmailCounts[email] = (toEmailCounts[email] || 0) + 1;
        });
        
        displayCcEmails.forEach(email => {
            ccEmailCounts[email] = (ccEmailCounts[email] || 0) + 1;
        });
        
        recipients.forEach(recipient => {
            let emailKey = recipient.email.toLowerCase();
            
            if (ccEmailCounts[emailKey] && ccEmailCounts[emailKey] > 0) {
                recipient.recipientType = RECIPIENT_TYPE_CC;
                ccEmailCounts[emailKey]--;
            }
            else if (toEmailCounts[emailKey] && toEmailCounts[emailKey] > 0) {
                recipient.recipientType = RECIPIENT_TYPE_TO;
                toEmailCounts[emailKey]--;
            }
        });
    }
    
    this.properties['recipients'] = { id: 0, value: recipients };
};

MsgReaderParser.prototype.getFieldValue = function(name) {
    let id = { subject: PROP_ID_SUBJECT, body: PROP_ID_BODY, bodyHTML: PROP_ID_HTML_BODY }[name];
    if (name === 'recipients') return this.properties['recipients'] ? this.properties['recipients'].value : [];
    return (this.properties[id]) ? this.properties[id].value : null;
};

// --- Exported Object ---
const MsgReader = {
    read: function(arrayBuffer) {
        let reader = new MsgReaderParser(arrayBuffer);
        if (reader.dataView.byteLength < 8) return reader.parseMime();
        let sig = reader.dataView.getUint32(0, true);
        return (sig === 0xE011CFD0) ? reader.parse() : reader.parseMime();
    }
};

export { MsgReader };