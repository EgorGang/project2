def handle_exceptions(func):
    """
    Декоратор для обработки исключений.
    """
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            args[0].chat.id
            bot = args[0].bot
            bot.send_message(args[0].chat.id, f"Ошибка: {str(e)}")
    return wrapper
