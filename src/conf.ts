import { AuthConfig } from 'node-sp-auth-config';

(new AuthConfig()).getContext()
  .catch(console.error);
