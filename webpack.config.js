const path = require('path');
const TerserPlugin = require('terser-webpack-plugin');
const webpack = require('webpack');

module.exports = {
  mode: "production",
  target: "web",
  entry: {
    'import-csv-discussions-action' : './src/import-csv-discussions-action.ts',
    'file-upload-dialog' : './src/file-upload-dialog.ts'
  },
  optimization: {
    minimizer: [
      new TerserPlugin({
        terserOptions: {
          output: {
            comments: false,
          },
        },
      }),
    ],
  },
  module: {
    rules: [
      {
        test: /\.ts?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
      {
          test: /\.(png|jpe?g|gif)$/i,
          loader: "file-loader",
          options: {
            outputPath: '../images',
          },
      }
    ],
  },
  resolve: {
    extensions: [ '.html', '.ts', '.js' ],
    fallback: {
      "buffer": require.resolve("buffer/")
    }
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
  },
  plugins: [
    new webpack.ProvidePlugin({
      Buffer: ['buffer', 'Buffer'],
  })
	]
};