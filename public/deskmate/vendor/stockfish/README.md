# Stockfish 11 (WebAssembly)

`stockfish.js` + `stockfish.wasm` are [Stockfish](https://stockfishchess.org) 11 compiled to
WebAssembly by [stockfish.js](https://github.com/nmrugg/stockfish.js) (npm `stockfish@11.0.0`),
unmodified. Deskmate runs it in a Web Worker for post-game analysis; it is a separate program
that Deskmate talks to over the UCI protocol.

License: GPLv3 (see `COPYING.txt`). Source code: https://github.com/nmrugg/stockfish.js
(tag/version 11.0.0) and https://github.com/official-stockfish/Stockfish (sf_11).
