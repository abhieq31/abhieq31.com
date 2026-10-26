// Things I've built, newest first. Shown on /projects.
export interface Project {
  name: string;
  tagline: string;
  why: string;
  live?: { href: string; label: string };
  code: string;
  stack: string;
}

export const PROJECTS: Project[] = [
  {
    name: "Deskmate",
    tagline: "Chess with the colleague at the next desk. One short code, no account.",
    why: "My company's Zscaler filter blocked chess.com and most other sites, and a few of us still wanted a game at lunch. So I wrote one that runs on any work laptop: a single HTML file, no sign-up, moves relayed peer to peer, Stockfish review after the game that even works when IT has switched WebAssembly off.",
    live: { href: "/deskmate", label: "Play" },
    code: "https://github.com/abhieq31/Deskmate",
    stack: "TypeScript · SCSS · Nostr relays · Stockfish 11",
  },
  {
    name: "Pragati",
    tagline: "Project and task management where everyone sees the whole board.",
    why: "Most status meetings exist because a contributor only sees their own slice of the work. I built Pragati for my teams in pharma quality, where every change must be recorded and nothing quietly deleted. It now runs in production for 5 teams and about 40 people.",
    live: { href: "https://pragatialm.vercel.app", label: "Demo" },
    code: "https://github.com/abhieq31/Pragati",
    stack: "Next.js · TypeScript · MongoDB",
  },
  {
    name: "Room to Breathe",
    tagline: "Open-source crowd-crush prevention for temples, stadiums and rallies.",
    why: "People in a crowd crush die because they cannot breathe, and every cause (too many people, too few exits, decisions made too late) is measurable before the event. India keeps losing people to stampedes, so I built a free tool any organiser can run: safe capacity, live density from CCTV, and a warning minutes before a zone turns critical.",
    live: { href: "https://abhipatelz.github.io/roomtobreathe/", label: "Open app" },
    code: "https://github.com/abhieq31/roomtobreathe",
    stack: "Python · computer vision · runs in the browser",
  },
];
