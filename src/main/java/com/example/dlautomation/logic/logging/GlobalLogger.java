package com.example.dlautomation.logic.logging;

import java.io.IOException;
import java.util.logging.*;

public class GlobalLogger {

    private static final Logger logger = Logger.getLogger(GlobalLogger.class.getName());

    public static void initialize(String logFilePath) throws IOException {
        for (Handler handler : logger.getHandlers()) {
            logger.removeHandler(handler);
        }

        FileHandler fileHandler = new FileHandler(logFilePath, true);
        fileHandler.setFormatter(new SimpleFormatter());
        logger.addHandler(fileHandler);

        logger.setLevel(Level.ALL);

        Logger rootLogger = Logger.getLogger("");
        Handler[] handlers = rootLogger.getHandlers();
        if (handlers[0] instanceof ConsoleHandler) {
            rootLogger.removeHandler(handlers[0]);
        }
    }

    public static Logger getLogger() {
        return logger;
    }
}
