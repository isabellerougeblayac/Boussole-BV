import { defineConfig } from 'vite';

export default defineConfig({
  base: '/',
  build: {
    // Éviter import.meta.url pour le logo : la version portable utilise un script classique.
    assetsInlineLimit: (file) => file.endsWith('logo-afigese-fond-blanc-1536x561.png') ? true : undefined,
  },
});
