// ─────────────────────────────────────────────────────────────────────────────
// Site configuration — edit this one file to make the whole site yours.
// ─────────────────────────────────────────────────────────────────────────────

export const site = {
  // The name shown in the top-left logo and used across the site.
  name: 'Abhi',
  title: 'Abhi — Thinking out loud',
  description: 'Technology, software, and business — thinking out loud.',
  author: 'Abhi',

  // Your social profiles (username: abhieq31).
  socials: {
    x: 'https://x.com/abhieq31',
    instagram: 'https://instagram.com/abhieq31',
    youtube: 'https://youtube.com/@abhieq31',
    substack: 'https://abhieq31.substack.com',
    apple: 'https://podcasts.apple.com/us/podcast/abhi/id1896933970',
    spotify: 'https://open.spotify.com/show/033ztaq1OfKJkcveYsMWIQ',
  },

  // Set true once the podcast is actually live — turns the homepage "Podcast"
  // row and the "Get podcast" menu on audio posts back on.
  showPodcast: false,

  // Links shown in the "Podcast" row on the homepage and the "Get podcast" menu.
  podcastLinks: [
    { label: 'X', href: 'https://x.com/abhieq31' },
    { label: 'Apple', href: 'https://podcasts.apple.com/us/podcast/abhi/id1896933970' },
    { label: 'Spotify', href: 'https://open.spotify.com/show/033ztaq1OfKJkcveYsMWIQ' },
    { label: 'YouTube', href: 'https://youtube.com/@abhieq31' },
    { label: 'Substack', href: 'https://abhieq31.substack.com' },
    { label: 'Email', href: '/subscribe' },
  ],

  // Newsletter / subscribe form. Single source of truth for the subscribe
  // endpoint — both the header's inline form and /subscribe read this.
  //   Buttondown:  https://buttondown.com/api/emails/embed-subscribe/abhieq31
  //   ConvertKit:  https://app.convertkit.com/forms/<FORM_ID>/subscriptions
  //   Substack:    https://<you>.substack.com/api/v1/free
  //   Formspree:   https://formspree.io/f/<FORM_ID>
  // Leave '' while styling/testing; replace before collecting real emails.
  newsletter: {
    action: 'https://abhieq31.substack.com/api/v1/free',
  },
} as const;

export type Site = typeof site;
