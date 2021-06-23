const path = require("path");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");

module.exports = {
  entry: "./client/src/main.ts",
  module: {
    rules: [{
      test: /\.ts$/,
      use: 'ts-loader' // compile TypeScript
    }]
  },
  plugins: [
    new CleanWebpackPlugin({
      cleanOnceBeforeBuildPatterns: ["**/*.js", "**/*.js.map"],
    })
  ],
  output: {
    filename: "[name].bundle.js",
    path: path.resolve(__dirname, "./dist/"),
  },
  resolve: {
    extensions: [
      '.ts', '.js',
    ]
  }
};
