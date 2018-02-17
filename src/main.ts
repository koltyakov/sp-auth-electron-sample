import { app, BrowserWindow, session } from 'electron';
import { getAuth } from 'node-sp-auth';
import { AuthConfig } from 'node-sp-auth-config';

let mainWindow: BrowserWindow = null;

app.on('window-all-closed', () => {
  app.quit();
});

app.on('ready', () => {
  (new AuthConfig()).getContext()
    .then(conf => {
      return getAuth(conf.siteUrl, conf.authOptions)
      .then(async (auth) => {

        const url = conf.siteUrl + '/';

        const cookies = auth.headers.Cookie.split('; ')
          .map(c => {
            const index = c.indexOf('=');
            const name = c.substring(0, index);
            const value = c.substring(index + 1, c.length);
            return { url, name, value };
          });

        console.log(cookies);

        for (const cookie of cookies) {
          await setCookiePromise(session.defaultSession.cookies, cookie);
        }

        mainWindow = new BrowserWindow({ width: 1024, height: 768 });

        mainWindow.loadURL(conf.siteUrl);

        mainWindow.webContents.session.cookies.get({ url: conf.siteUrl }, (err, cookies) => {
          if (err) {
            return console.log(err);
          }
          console.log(cookies);
        });

      });
    })
    .catch(error => {
      console.log(error);
      app.quit();
    });
});

const setCookiePromise = (sessionCookies, cookie) => {
  return new Promise((resolve, reject) => {
    sessionCookies.set(cookie, error => {
      if (error) {
        return reject(error);
      }
      resolve();
    });
  });
};
