package com.example.csvexcelexport.exception;

public class ServiceExcepetion extends RuntimeException{

    private static final long serialVersionID = 1L;

    public ServiceExcepetion() { super(); }

    public ServiceExcepetion(String message) { super(message); }

    public ServiceExcepetion(Throwable cause) { super(cause); }
}
