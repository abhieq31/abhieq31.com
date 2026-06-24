import { SITE } from "./consts";

export const SOCIALS = [
  {
    name: "X",
    href: "https://x.com/abhieq31",
    linkTitle: `${SITE.title} on X`,
    icon: "twitter",
    active: true,
  },
  {
    name: "Instagram",
    href: "https://instagram.com/abhieq31",
    linkTitle: `${SITE.title} on Instagram`,
    icon: "instagram",
    active: true,
  },
  {
    name: "Github",
    href: "https://github.com/abhieq31",
    linkTitle: `${SITE.title} on GitHub`,
    icon: "github",
    active: true,
  },
] as const;
