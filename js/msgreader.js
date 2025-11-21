/**
 * Email File Parser
 * Decodes .msg, .oft, and .eml files.
 * Handles OLE compound documents, MIME decoding, and property extraction.
 */

'use strict';

// MAPI Property Tags
const PROP_TYPE_STRING = 0x001E;
const PROP_TYPE_STRING8 = 0x001F;
const PROP_ID_SUBJECT = 0x0037;
const PROP_ID_BODY = 0x1000;
const PROP_ID_HTML_BODY = 0x1013;
const PROP_ID_DISPLAY_TO = 0x0E04;
const PROP_ID_DISPLAY_CC = 0x0E03;
const PROP_ID_DISPLAY_BCC = 0x0E02;

const RECIPIENT_TYPE_TO = 1;
const RECIPIENT_TYPE_CC = 2;
const RECIPIENT_TYPE_BCC = 3;

// Decoders
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

function _decodeQuotedPrintable(str, charset = 'utf-8') {
    if (!str) return '';
    let decoded = str
        .replace(/=(\r\n|\n)/g, '')
        .replace(/=([0-9A-F]{2})/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
    
    try {
        let bytes = new Uint8Array(decoded.length);
        for (let i = 0; i < decoded.length; i++) bytes[i] = decoded.charCodeAt(i);
        let encoding = charset.toLowerCase() === 'us-ascii' ? 'utf-8' : charset.toLowerCase();
        return new TextDecoder(encoding, { fatal: false }).decode(bytes);
    } catch (e) { 
        return decoded; 
    }
}

function _stripHtml(html) {
    if (!html || typeof html !== 'string') return ''; 
    let text = html
        .replace(/<head[\s\S]*?<\/head>/gi, '')
        .replace(/<style[\s\S]*?<\/style>/gi, '')
        .replace(/<script[\s\S]*?<\/script>/gi, '')
        .replace(/<!--[\s\S]*?-->/g, '')
        .replace(/[.#a-z0-9_]+\s*\{[^}]+\}/gi, '');

    text = text.replace(/\s*<(br|p|div|tr|li|h1|h2|h3|h4|h5|h6)[^>]*>\s*/gi, '\n');
    
    let parser = getDOMParser();
    if (parser) {
        try {
            let doc = parser.parseFromString(text, 'text/html');
            doc.querySelectorAll('style, script, link, meta, title').forEach(el => el.remove());
            text = doc.body ? doc.body.textContent : (doc.documentElement.textContent || '');
        } catch (e) {
            text = text.replace(/<[^>]+>/g, '');
        }
    } else {
        text = text.replace(/<[^>]+>/g, '');
    }
    
    return _normalizeText(text);
}

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
            let decoded = new TextDecoder('utf-8', { fatal: true }).decode(view);
            return decoded.split('\0')[0];
        } catch (e) { return dataViewToString(view, 'ascii'); }
    }
    if (encoding === 'utf16le') {
        try {
            let decoded = getTextDecoder('utf-16le').decode(view);
            return decoded.split('\0')[0];
        } catch (e) { return ''; }
    }
    try {
        let decoded = getTextDecoder('windows-1252').decode(view);
        return decoded.split('\0')[0];
    } catch(e) { return ''; }
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

// Parser Logic
function MsgReaderParser(arrayBuffer) {
    this.buffer = arrayBuffer instanceof ArrayBuffer ? arrayBuffer : new Uint8Array(arrayBuffer).buffer;
    this.dataView = new DataView(this.buffer);
    this.header = null; this.fat = null; this.miniFat = null;
    this.directoryEntries = []; this.properties = {};
}

MsgReaderParser.prototype.parse = function() {
    if (this.dataView.getUint32(0, true) !== 0xE011CFD0) return this.parseMime();

    this.readHeader(); 
    this.readFAT(); 
    this.readMiniFAT(); 
    this.readDirectory(); 
    this.extractProperties();

    return {
        subject: this.getFieldValue('subject'),
        body: this.getFieldValue('body'),
        bodyHTML: this.getFieldValue('bodyHTML'),
        recipients: this.getFieldValue('recipients')
    };
};

MsgReaderParser.prototype.parseMime = function() {
    let rawText = '';
    try { rawText = new TextDecoder('utf-8', { fatal: false }).decode(this.dataView); }
    catch (e) { 
        try { rawText = new TextDecoder('latin1').decode(this.dataView); } catch (e2) { rawText = ''; }
    }
    
    let result = { subject: null, to: null, cc: null, body: null };
    
    const findField = (name) => {
        const search = new RegExp(`\\b${name}:\\s*([^\\r\\n]+)`, 'i');
        const match = rawText.match(search);
        return match ? match[1].trim() : null;
    };
    
    result.subject = findField('Subject');
    let recipients = [];
    
    const parseMimeAddresses = (str, type) => {
        if (!str) return;
        str.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).forEach(addr => {
            let parsed = parseAddress(addr);
            if (parsed.email) recipients.push({ name: parsed.name, email: parsed.email, recipientType: type });
        });
    };

    parseMimeAddresses(findField('To'), RECIPIENT_TYPE_TO);
    parseMimeAddresses(findField('Cc'), RECIPIENT_TYPE_CC);
    let bccMatch = rawText.match(/^Bcc:\s*([^\r\n]+)/im);
    if (bccMatch) parseMimeAddresses(bccMatch[1].trim(), RECIPIENT_TYPE_BCC);

    let headerEnd = rawText.indexOf('\r\n\r\n');
    if (headerEnd === -1) headerEnd = rawText.indexOf('\n\n');
    
    if (headerEnd !== -1) {
        let bodyText = rawText.substring(headerEnd + 4);
        let charset = 'utf-8';
        let encoding = null;

        const charsetMatch = rawText.match(/charset=["']?([^"';\r\n]+)/i);
        if (charsetMatch) charset = charsetMatch[1];

        if (rawText.match(/^Content-Transfer-Encoding:\s*quoted-printable/im)) encoding = 'quoted-printable';
        
        result.body = (encoding === 'quoted-printable') ? _decodeQuotedPrintable(bodyText, charset) : bodyText;
    }

    return {
        subject: result.subject,
        body: result.body,
        bodyHTML: null,
        recipients: recipients
    };
};

MsgReaderParser.prototype.readHeader = function() {
    this.header = {
        sectorShift: this.dataView.getUint16(30, true),
        miniSectorShift: this.dataView.getUint16(32, true),
        fatSectors: this.dataView.getUint32(44, true),
        directoryFirstSector: this.dataView.getUint32(48, true),
        miniFatFirstSector: this.dataView.getUint32(60, true),
        miniFatTotalSectors: this.dataView.getUint32(64, true),
    };
    this.header.sectorSize = Math.pow(2, this.header.sectorShift);
    this.header.miniSectorSize = Math.pow(2, this.header.miniSectorShift);
};

MsgReaderParser.prototype.readFAT = function() {
    let sectorSize = this.header.sectorSize;
    let entries = sectorSize / 4;
    this.fat = [];
    for (let i = 0; i < this.header.fatSectors; i++) {
        let s = this.dataView.getUint32(76 + i * 4, true);
        if (s !== 0xFFFFFFFE && s !== 0xFFFFFFFF) {
            let offset = 512 + s * sectorSize;
            for (let j = 0; j < entries; j++) {
                this.fat.push(this.dataView.getUint32(offset + j * 4, true));
            }
        }
    }
};

MsgReaderParser.prototype.readMiniFAT = function() {
    if (this.header.miniFatFirstSector === 0xFFFFFFFE) { this.miniFat = []; return; }
    this.miniFat = [];
    let sector = this.header.miniFatFirstSector;
    while (sector !== 0xFFFFFFFE && sector < this.fat.length) {
        let offset = 512 + sector * this.header.sectorSize;
        for (let i = 0; i < this.header.sectorSize / 4; i++) {
            this.miniFat.push(this.dataView.getUint32(offset + i * 4, true));
        }
        sector = this.fat[sector];
    }
};

MsgReaderParser.prototype.readDirectory = function() {
    let sector = this.header.directoryFirstSector;
    while (sector !== 0xFFFFFFFE) {
        let offset = 512 + sector * this.header.sectorSize;
        for (let i = 0; i < this.header.sectorSize / 128; i++) {
            let entry = this.readDirectoryEntry(offset + i * 128);
            if (entry) this.directoryEntries.push(entry);
        }
        sector = this.fat[sector];
    }
    this.directoryEntries.forEach((de, idx) => de.id = idx);
};

MsgReaderParser.prototype.readDirectoryEntry = function(offset) {
    let nameLen = this.dataView.getUint16(offset + 64, true);
    if (nameLen > 64) nameLen = 64;
    let name = dataViewToString(new DataView(this.buffer, offset, nameLen), 'utf16le');
    let type = this.dataView.getUint8(offset + 66);
    return {
        name: name, type: type,
        startSector: this.dataView.getUint32(offset + 116, true),
        size: this.dataView.getUint32(offset + 120, true),
        childId: this.dataView.getInt32(offset + 76, true)
    };
};

MsgReaderParser.prototype.readStream = function(entry) {
    if (!entry || entry.size === 0) return new Uint8Array(0);
    let isMini = entry.size < 4096;
    let chain = isMini ? this.miniFat : this.fat;
    let startSector = entry.startSector;
    let sectorSize = isMini ? this.header.miniSectorSize : this.header.sectorSize;
    let result = new Uint8Array(entry.size);
    let offset = 0;
    
    let root = this.directoryEntries[0];
    let miniStream = isMini ? this.readStream({ ...root, size: root.size, startSector: root.startSector }) : null;

    while (startSector !== 0xFFFFFFFE && offset < entry.size) {
        let copySize = Math.min(sectorSize, entry.size - offset);
        if (isMini) {
            result.set(miniStream.slice(startSector * sectorSize, (startSector + 1) * sectorSize).slice(0, copySize), offset);
        } else {
            let sectorOffset = 512 + startSector * sectorSize;
            result.set(new Uint8Array(this.buffer, sectorOffset, copySize), offset);
        }
        offset += copySize;
        startSector = chain[startSector];
    }
    return result;
};

MsgReaderParser.prototype.extractProperties = function() {
    let rawProps = {};
    this.directoryEntries.forEach(entry => {
        if (entry.name.indexOf('__substg1.0_') !== 0) return;
        let propTag = parseInt(entry.name.substring(12, 16), 16);
        rawProps[propTag] = this.readStream(entry);
    });

    const getStr = (id) => {
        let data = rawProps[id];
        if (!data) return null;
        let view = new DataView(data.buffer, data.byteOffset, data.byteLength);
        let s = dataViewToString(view, 'utf16le');
        if (!s || s.length < 2) s = dataViewToString(view, 'utf-8');
        return s;
    };

    let body = getStr(PROP_ID_BODY);
    let htmlBody = getStr(PROP_ID_HTML_BODY);
    
    if (!body && htmlBody) body = _stripHtml(htmlBody);
    
    this.properties[PROP_ID_SUBJECT] = { value: getStr(PROP_ID_SUBJECT) };
    this.properties[PROP_ID_BODY] = { value: _normalizeText(body) };
    this.properties[PROP_ID_HTML_BODY] = { value: htmlBody };

    let recipients = [];
    let displayTo = getStr(PROP_ID_DISPLAY_TO);
    let displayCc = getStr(PROP_ID_DISPLAY_CC);
    
    if (displayTo) {
        displayTo.split(';').forEach(e => {
            let p = parseAddress(e);
            if (p.email) recipients.push({ email: p.email, recipientType: RECIPIENT_TYPE_TO });
        });
    }
    if (displayCc) {
        displayCc.split(';').forEach(e => {
            let p = parseAddress(e);
            if (p.email) recipients.push({ email: p.email, recipientType: RECIPIENT_TYPE_CC });
        });
    }

    this.properties['recipients'] = { value: recipients };
};

MsgReaderParser.prototype.getFieldValue = function(name) {
    if (name === 'recipients') return this.properties['recipients'] ? this.properties['recipients'].value : [];
    let map = { subject: PROP_ID_SUBJECT, body: PROP_ID_BODY, bodyHTML: PROP_ID_HTML_BODY };
    return this.properties[map[name]] ? this.properties[map[name]].value : null;
};

const MsgReader = {
    read: function(arrayBuffer) {
        return new MsgReaderParser(arrayBuffer).parse();
    }
};

export { MsgReader };
