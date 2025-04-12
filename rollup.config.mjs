import commonjs from '@rollup/plugin-commonjs';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import cleanup from 'rollup-plugin-cleanup';
import nodePolyfills from 'rollup-plugin-polyfill-node';
import prettier from 'rollup-plugin-prettier';
import typescript from 'rollup-plugin-typescript2';

function preventTreeShakingPlugin() {
  return {
    name: 'no-treeshaking',
    resolveId(id, importer) {
      if (!importer) {
        // don't treeshake entry points
        return { id, moduleSideEffects: 'no-treeshake' };
      }
      return null;
    },
  };
}

export default {
  input: 'src/index.ts',
  output: {
    dir: 'dist',
    format: 'esm',
  },
  plugins: [
    preventTreeShakingPlugin(),
    nodeResolve(),
    nodePolyfills({ buffer: true }),
    commonjs(),
    typescript(),
    cleanup({ comments: 'none', extensions: ['.ts'] }),
    prettier({ parser: 'typescript' }),
  ],
  context: 'this',
};
