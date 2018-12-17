# YT Tracker

[![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg)](https://github.com/google/clasp)

Easily fetch statistics like number of subscribers, likes, dislikes, comments of any YouTube videos/channels using [YT Tracker](https://gsuite.google.com/marketplace/app/yt_tracker/952783286913).

## How to install & configure

- As this add-on is built using [clasp](https://github.com/google/clasp), install it globally on your machine.

```
npm install @google/clasp -g
clasp login // to login to your Google account for authentication
```

- Now clone this repo and install dependencies.

```
git clone https://github.com/vicke4/yt-tracker && cd yt-tracker
npm install
```

- Edit .clasp.json and add your script id in the json.

- Run `npm run build` to generate bundled code that'll be pushed to your script. Then run clasp push.

We thank [labnol](https://github.com/labnol) for the boilerplate - [apps-script-starter](https://github.com/labnol/apps-script-starter).

## License

MIT
