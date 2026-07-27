// Place any global data in this file.
// You can import this data from anywhere in your site by using the `import` keyword.

interface SocialLink {
  href: string;
  label: string;
}

interface Site {
  website: string;
  author: string;
  profile: string;
  desc: string;
  title: string;
  ogImage: string;
  lightAndDarkMode: boolean;
  postPerIndex: number;
  postPerPage: number;
  scheduledPostMargin: number;
  showArchives: boolean;
  showBackButton: boolean;
  editPost: {
    enabled: boolean;
    text: string;
    url: string;
  };
  dynamicOgImage: boolean;
  lang: string;
  timezone: string;
}

// Site configuration
export const SITE: Site = {
  website: "https://abhieq31.com/",
  author: "Abhi",
  profile: "https://abhieq31.com/about",
  desc: "Technology, software, and business — thinking out loud.",
  title: "Abhi",
  ogImage: "icon-512.png",
  lightAndDarkMode: false, // nav.al-style: always dark
  postPerIndex: 10,
  postPerPage: 10,
  scheduledPostMargin: 15 * 60 * 1000,
  showArchives: true,
  showBackButton: false,
  editPost: {
    enabled: false,
    text: "Edit on GitHub",
    url: "https://github.com/abhieq31/abhieq31.com/edit/main/",
  },
  dynamicOgImage: true,
  lang: "en",
  timezone: "America/Los_Angeles",
};

export const SITE_TITLE = SITE.title;
export const SITE_DESCRIPTION = SITE.desc;

// Navigation links (nav.al pattern: Archive + socials + Subscribe in Header)
export const NAV_LINKS: SocialLink[] = [
  {
    href: "/posts",
    label: "Archive",
  },
  {
    href: "https://x.com/abhieq31",
    label: "X",
  },
  {
    href: "https://www.instagram.com/abhieq31/",
    label: "Instagram",
  },
  {
    href: "https://github.com/abhieq31",
    label: "GitHub",
  },
];

// Social media links
export const SOCIAL_LINKS: SocialLink[] = [
  {
    href: "https://x.com/abhieq31",
    label: "X",
  },
  {
    href: "https://www.instagram.com/abhieq31/",
    label: "Instagram",
  },
  {
    href: "https://github.com/abhieq31",
    label: "GitHub",
  },
  {
    href: "/rss.xml",
    label: "RSS",
  },
];

// Icon map for social media
export const ICON_MAP: Record<string, string> = {
  X: "twitter",
  GitHub: "github",
  RSS: "rss",
  Email: "mail",
};
