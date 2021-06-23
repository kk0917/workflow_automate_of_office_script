const { merge } = require("webpack-merge");
const common = require("./webpack.common.js");

const demoData = require("./httpdocs/demo_data.json");

module.exports = merge(common, {
  mode: "development",
  devtool: "inline-source-map",
  devServer: {
    contentBase: "./httpdocs/",
    inline: true,
    hot: true,
    liveReload: true,
    compress: true,
    host: "0.0.0.0",
    port: 8080,
    before: (app) => {
      app.post("/path/to/end-point", (req, res) => {
        res.json(demoData.eventID);
      });
      app.post("/path/to/end-point", (req, res) => {
        res.json(demoData.FBP);
      });
    }
  }
});
