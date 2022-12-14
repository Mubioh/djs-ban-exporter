const { SlashCommandBuilder, REST, Routes } = require('discord.js');
require('dotenv').config();

const Excel = require('exceljs');

const rest = new REST({ version: '10' }).setToken(process.env.DISCORD_TOKEN);

module.exports = {
	data: new SlashCommandBuilder()
		.setName('export')
		.setDescription('Exports an Excel file with all banned users.'),
	async execute(interaction) {
		const results = [];
		const bans = await rest.get(Routes.guildBans(process.env.GUILD_ID));

		for (const ban of bans) {
			let row = [];

			row.push(ban.user.username + '#' + ban.user.discriminator);
			row.push(ban.user.id);
			if (ban.reason) {
				row.push(ban.reason);
			}

			results.push(row);
		}

		const workbook = new Excel.Workbook();
		const worksheet = workbook.addWorksheet(new Date().toISOString().slice(0, 10));
		worksheet.columns = [
			{ header: 'Username', key: 'username ', width: 30 },
			{ header: 'User ID', key: 'userID', width: 30 },
			{ header: 'Reason', key: 'reason', width: 50 },
		];

		for (const result of results) {
			worksheet.addRow(result);
		}

		const filePath =
			'./data/' + interaction.guild.name.toLowerCase().replace(/ /g, '-').replace(/[^a-z0-9-]/g, '') +'.xlsx';
		await workbook.xlsx.writeFile(filePath);

		await interaction.reply({
			content: 'Successfully compiled the list of banned users:',
			files: [filePath],
			ephemeral: true,
		});
	},
};
