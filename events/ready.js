const { Events, REST, Routes } = require('discord.js');
require('dotenv').config();

module.exports = {
	name: Events.ClientReady,
	once: true,
	async execute(client) {
		const rest = new REST({ version: '10' }).setToken(process.env.DISCORD_TOKEN);
		const commands = Array.from(client.commands.values()).map((map) => map.data);

		try {
			console.log(`Started refreshing ${commands.length} application (/) commands.`);
			const data = await rest.put(Routes.applicationCommands(client.user.id), {
				body: commands,
			});
			console.log(`Successfully reloaded ${data.length} application (/) commands.`);
		} catch (error) {
			console.error(error);
		}

		console.log(`Ready! Logged in as ${client.user.tag} (${client.user.id})`);
	},
};
