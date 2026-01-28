const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isDevelopment = argv.mode === 'development';
  const isProduction = env && env.production === true;

  // Output to /docs for production (GitHub Pages), /dist for development
  const outputPath = isProduction
    ? path.resolve(__dirname, 'docs')
    : path.resolve(__dirname, 'dist');

  // Base URL for assets - GitHub Pages URL for production, localhost for dev
  const publicPath = isProduction
    ? 'https://mikedtcm22.github.io/excel_conditional_fill/'
    : 'https://localhost:3000/';

  // Copy patterns for assets
  const copyPatterns = [
    {
      from: 'assets',
      to: 'assets',
      noErrorOnMissing: true
    }
  ];

  // For production, include shortcuts.json in the build output
  if (isProduction) {
    copyPatterns.push({
      from: 'src/shortcuts.json',
      to: 'shortcuts.json',
      noErrorOnMissing: true
    });
    copyPatterns.push({
      from: 'manifest-production.xml',
      to: 'manifest-production.xml',
      noErrorOnMissing: true
    });
  } else {
    // For development, copy manifest.xml and shortcuts.json from src
    copyPatterns.push({
      from: 'manifest.xml',
      to: 'manifest.xml',
      noErrorOnMissing: true
    });
    copyPatterns.push({
      from: 'src/shortcuts.json',
      to: 'shortcuts.json',
      noErrorOnMissing: true
    });
  }

  return {
    mode: isDevelopment ? 'development' : 'production',
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts'
    },
    output: {
      filename: '[name].js',
      path: outputPath,
      clean: true,
      publicPath: isProduction ? './' : '/'
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
        patterns: copyPatterns
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
