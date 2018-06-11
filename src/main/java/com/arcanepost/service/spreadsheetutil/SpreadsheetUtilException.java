package com.arcanepost.service.spreadsheetutil;

/**
 * Exception used throughout the utility package.
 *
 * @author callen
 * @since 01.25.2016
 */
public class SpreadsheetUtilException extends Exception {

    private static final long serialVersionUID = -5214028645004125161L;

    /**
     * Empty Constructor.
     */
    public SpreadsheetUtilException() {
    }

    /**
     * Constructor. Pass param to super.
     *
     * @param e
     *            The exception to wrap.
     */
    public SpreadsheetUtilException(Exception e) {
        super(e);
    }

    /**
     * Constructor. Pass param to super.
     *
     * @param message
     *            The message to wrap.
     */
    public SpreadsheetUtilException(String message) {
        super(message);
    }
}
