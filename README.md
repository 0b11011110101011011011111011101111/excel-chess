# Chess... for Excel!

A basic chess engine that runs entirely in Excel using `VBA`. It only looks a few moves ahead so it won't be beating `StockFish` anytime soon, but it can catch you unawares if you aren't paying attention.

### Installation

Make sure the `Chess_Pieces/` folder is located in the same directory as `Chess.xlsm`, so that the visual resources are available.

This file has the `.xlsm` extension, indicating that it utilizes `VBA` code (referred to as `macros` by Excel). These extensions will often trigger antivirus warnings when trying to download. I promise there's nothing suspicious about the contents of these files.

### Features

- Graphical display for chess pieces
- Allows for default initial position, or randomly-generated `Chess960` position
- Automatically loads any board position from a pasted [`FEN`](https://en.wikipedia.org/wiki/Forsyth%E2%80%93Edwards_Notation), and generates `FEN` strings from every board position
- Allows for player input via [Algebraic Notation](https://en.wikipedia.org/wiki/Algebraic_notation_(chess))
- Allows for computer turns, usually taking 5-10 seconds to look ahead ~3 turns, and evaluating positions based on a simple minimax algorithm
- Automatically records [`PGN`](https://en.wikipedia.org/wiki/Portable_Game_Notation) for ongoing game, and stores `FEN` strings for every ply

### Areas for Improvement

- There's very little error catching in the player move input. Incorrect move inputs can sometimes yield illegal or crazy moves.
- The interface allows either player or computer to make a move an anytime. This gives flexibility, but it's left to the user to keep track of turns to play an actual game.
- Not very good catching for endgame scenarios. A textbox will appear if the computer tries to move in a checkmate or stalemate position, but it will not recognize when a player or computer has been put into this position.
