// ─────────────────────────────────────────────────────────────────────────────
// Site configuration — edit this one file to make the whole site yours.
// ─────────────────────────────────────────────────────────────────────────────

export const site = {
  // The name shown in the top-left logo and used across the site.
  name: 'Abhi',
  title: 'Abhi',
  description: 'Personal site of Abhi — essays, notes, and a podcast.',
  author: 'Abhi',

  // Your social profiles (username: abhieq31).
  socials: {
    x: 'https://x.com/abhieq31',
    instagram: 'https://instagram.com/abhieq31',
    youtube: 'https://youtube.com/@abhieq31',
  },

  // Links shown in the "Podcast" row on the homepage and the "Get podcast" menu.
  // Replace '#' with your real show URLs as you publish on each platform.
  podcastLinks: [
    { label: 'X', href: 'https://x.com/abhieq31' },
    { label: 'Apple', href: '#' },
    { label: 'Spotify', href: '#' },
    { label: 'YouTube', href: 'https://youtube.com/@abhieq31' },
    { label: 'Substack', href: '#' },
    { label: 'Email', href: '/subscribe' },
  ],

  // Newsletter / subscribe form.
  // Paste your provider's form endpoint here to make the subscribe form live.
  //   Buttondown:  https://buttondown.com/api/emails/embed-subscribe/abhieq31
  //   ConvertKit:  https://app.convertkit.com/forms/<FORM_ID>/subscriptions
  //   Substack:    https://<you>.substack.com/api/v1/free  (or embed)
  //   Formspree:   https://formspree.io/f/<FORM_ID>
  // Leave '' to show a "not connected yet" notice instead.
  newsletter: {
    action: '',
  },
} as const;

export type Site = typeof site;
