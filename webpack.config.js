const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isDevelopment = argv.mode === 'development';

  return {
    mode: isDevelopment ? 'development' : 'production',
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts'
    },
    output: {
      filename: '[name].js',
      path: path.resolve(__dirname, 'dist'),
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.js']
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: 'ts-loader',
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'manifest.xml',
            to: 'manifest.xml',
            noErrorOnMissing: true
          },
          {
            from: 'assets',
            to: 'assets',
            noErrorOnMissing: true
          }
        ]
      })
    ],
    devtool: isDevelopment ? 'source-map' : false,
    devServer: {
      static: {
        directory: path.join(__dirname, 'dist')
      },
      port: 3000,
      hot: true,
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      server: 'https',
      client: {
        overlay: {
          errors: true,
          warnings: false
        }
      }
    }
  };
};
