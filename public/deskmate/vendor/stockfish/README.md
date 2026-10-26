# Stockfish 11 (WebAssembly)

`stockfish.js` + `stockfish.wasm` (and the plain-JavaScript fallback `stockfish.asm.js`) are [Stockfish](https://stockfishchess.org) 11 compiled to
WebAssembly by [stockfish.js](https://github.com/nmrugg/stockfish.js) (npm `stockfish@11.0.0`),
unmodified. Deskmate runs it in a Web Worker for post-game analysis; it is a separate program
that Deskmate talks to over the UCI protocol.

License: GPLv3 (see `COPYING.txt`). Source code: https://github.com/nmrugg/stockfish.js
(tag/version 11.0.0) and https://github.com/official-stockfish/Stockfish (sf_11).

`stockfish.asm.js` is used when the browser has no WebAssembly, e.g. Chrome with the JIT
switched off by company policy. It is about 25 times slower, so analysis then searches less deep.
