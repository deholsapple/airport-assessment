// Manually researched FBO directory, built up on demand (no bulk/authoritative
// government source exists for FBO data, unlike airport-data.js). Each entry
// is looked up individually via web search or the AirNav link in Section 6 --
// verify against current sources before relying on it operationally.
//
// Structure: window.FBO_DATA[ICAO] = [ { name, phone, fuel, gpu, deice, hangar, hours, sourceNote } ]
window.FBO_DATA = {};
