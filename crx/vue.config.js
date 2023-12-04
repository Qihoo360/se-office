const { defineConfig } = require('@vue/cli-service')
const { ElementPlusResolver } = require('unplugin-vue-components/resolvers')
const CopyWebpackPlugin = require('copy-webpack-plugin')
const { CleanWebpackPlugin } = require('clean-webpack-plugin')
const CrxPlugin = require('webpack-crx')
const path = require('path')
const pkg = require('./package.json')

const isDev = process.env.NODE_ENV === 'development'

const PATH = {
  output: path.resolve('dist'),
  build: path.resolve('build'),
}

// Generate pages object
const pagesObj = {}
const chromeExtPages = ['background', 'app', 'popup']

chromeExtPages.forEach(name => {
  pagesObj[name] = {
    entry: `src/${name}.ts`,
    template: `src/${name}.html`,
    filename: `${name}.html`,
  }
})

const manifest = {
  from: 'src/manifest.json',
  to: 'manifest.json',
  transform(content) {
    if (!isDev) {
      return content
    }

    let manifest = JSON.parse(content.toString())
    if (manifest.background && manifest.background.scripts) {
      manifest.background.scripts.unshift('hot-reload.js')
    }

    return JSON.stringify(manifest)
  },
}

const copyImgs = {
  from: 'src/img/',
  to: `img/`,
}

const copyOffice = [
  {
    from: '../web-apps/deploy/sdkjs/',
    to: './sdkjs/',
  },
  {
    from: '../web-apps/deploy/web-apps/',
    to: './web-apps/',
  },
  {
    from: './plugin',
    to: './plugin/',
  },
]

const plugins = [
  require('unplugin-auto-import/webpack')({
    resolvers: [ElementPlusResolver()],
    // eslintrc: {
    //   enabled: false, // Default `false`
    //   filepath: './.eslintrc-auto-import.json',
    //   globalsPropValue: true,
    // },
  }),
  require('unplugin-vue-components/webpack')({
    resolvers: [ElementPlusResolver()],
  }),
  new CopyWebpackPlugin([manifest, copyImgs, ...copyOffice]),
]

if (isDev) {
  plugins.push(
    new CopyWebpackPlugin([
      {
        from: 'src/utils/hot-reload.js',
        to: PATH.output,
      },
    ])
  )
}

const debugMode = process.argv[5] === 'debug'
if (!isDev || debugMode) {
  if (!debugMode) {
    plugins.push(
      new CleanWebpackPlugin({
        cleanOnceBeforeBuildPatterns: [PATH.build],
      })
    )
  }
}

module.exports = defineConfig({
  transpileDependencies: true,
  pages: pagesObj,
  productionSourceMap: false,
  // publicPath: '.',

  configureWebpack: {
    entry: {
      content: './src/content/content.ts',
    },
    output: {
      filename: 'js/[name].js',
    },
    plugins: plugins,
    resolve: {
      alias: {
        '@': path.resolve(__dirname, 'src/'),
      },
    },
  },
  css: {
    extract: {
      filename: 'css/[name].css',
      // chunkFilename: 'css/[name].css'
    },
    loaderOptions: {
      sass: {},
    },
  },

  chainWebpack: config => {
    if (isDev) {
      config.devtool = 'eval-source-map'
    }
    config.optimization.splitChunks(false)

    config.optimization.minimizer('terser').tap(args => {
      args[0].terserOptions.format = { ascii_only: true }
      return args
    })

    const fontsRule = config.module.rule('fonts')

    fontsRule.uses.clear()
    fontsRule
      .test(/\.(woff2?|eot|ttf|otf)(\?.*)?$/i)
      .use('url')
      .loader('url-loader')
      .options({
        limit: 2000,
        name: 'fonts/[name].[ext]',
      })
    if (process.env.npm_config_report) {
      config
        .plugin('webpack-bundle-analyzer')
        .use(require('webpack-bundle-analyzer').BundleAnalyzerPlugin)
    }
  },
})
