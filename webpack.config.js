const devCerts = require('office-addin-dev-certs');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const DotenvPlugin = require('dotenv-webpack');
const path = require('path');
require('dotenv').config();

module.exports = async (env, options) => {
	return {
		output: {
			path: path.resolve(__dirname, 'dist'),
		},
		target: ['web', 'es5'],
		devtool: 'source-map',
		entry: {
			commands: './src/commands/commands.ts',
		},
		resolve: {
			extensions: ['.ts', '.tsx', '.html', '.js'],
		},
		module: {
			rules: [
				{
					test: /\.ts$/,
					exclude: /node_modules/,
					use: 'ts-loader',
				},
				{
					test: /\.html$/,
					exclude: /node_modules/,
					use: 'html-loader',
				},
			],
		},
		plugins: [
			new CleanWebpackPlugin(),
			new HtmlWebpackPlugin({
				filename: 'commands.html',
				template: './src/commands/commands.html',
				chunks: ['polyfill', 'commands'],
				hash: true,
			}),
			new CopyWebpackPlugin({
				patterns: [
					{
						from: 'manifest.xml',
						to: 'manifest.xml',
					},
				].filter(Boolean),
			}),
			new DotenvPlugin(),
		],
		devServer: {
			headers: {
				'Access-Control-Allow-Origin': '*',
			},
			https:
				options.https !== undefined
					? options.https
					: await devCerts.getHttpsServerOptions().then((config) => {
							// Unsuported key.
							delete config.ca;
							return config;
					  }),
			port: process.env.npm_package_config_dev_server_port || 3000,
		},
		performance: {
			maxAssetSize: 512 << 10,
			maxEntrypointSize: 512 << 10,
		},
		optimization: {
			minimize: true,
		},
	};
};
