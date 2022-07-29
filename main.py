import os
import telebot
import logging
import openpyxl
from config import *
from flask import Flask, request
from telebot import types

bot = telebot.TeleBot(BOT_TOKEN)
server = Flask(__name__)
logger = telebot.logger
logger.setLevel(logging.DEBUG)

path = 'C:\Users\Asus\PycharmProjects\myplate\users.xlsx'
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
user_id = ""
username = ""
A = 0
name = ""
age = 0
height = 0
weight = 0
sex = ""


@bot.message_handler(commands=['start'])
def start(message):
    username = message.from_user.username
    bot.reply_to(message, f'Hi, {username}')


@server.route(f"/{BOT_TOKEN}", methods=["POST"])
def redirect_message():
    json_string = request.get_data().decode("utf-8")
    update = telebot.types.Update.de_json(json_string)
    bot.process_new_updates([update])
    return "!", 200


if __name__ == "__main__":
    bot.remove_webhook()
    bot.set_webhook(url=APP_URL)
    server.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))


