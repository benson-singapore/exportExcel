package com.jeff.regan.excel.util;

import java.util.function.*;

/**
 * Function 工具类，对lambda调用
 * @author  zhangby
 * @date  2017/8/5 17:31
 */
public class FunctionUtils {

    /**
     * 接收T对象并返回boolean
     * @param t
     * @param predicate
     * @param <T>
     * @return
     */
    public static <T> boolean predicate(T t,Predicate<T> predicate){
         return predicate.test(t);
    }

    /**
     * 接收T对象，返回E对象
     * @param t
     * @param func
     * @param <T>
     * @param <E>
     * @return
     */
    public static  <T,E> E function(T t,Function<T,E> func){
        return func.apply(t);
    }

    /**
     * 接收T对象，不返回值
     * @param t
     * @param consumer
     * @param <T>
     */
    public static <T> void consumer(T t ,Consumer<T> consumer) {
        consumer.accept(t);
    }

    /**
     * 提供T对象（例如工厂），不接收值
     * @param supplier
     * @param <T>
     * @return
     */
    public static <T> T supplier(Supplier<T> supplier) {
        return supplier.get();
    }

    /**
     * 接收T对象，返回T对象
     * @param t
     * @param supplier
     * @param <T>
     * @return
     */
    public static <T> T unaryOperator(T t,UnaryOperator<T> supplier) {
        return supplier.apply(t);
    }

    /**
     * 接收两个T对象，返回T对象
     * @param t1
     * @param t2
     * @param supplier
     * @param <T>
     * @return
     */
    public static <T> T binaryOperator(T t1,T t2,BinaryOperator<T> supplier) {
        return supplier.apply(t1,t2);
    }

    /**
     * 接收T对象和U对象，返回R对象
     * @param t
     * @param u
     * @param supplier
     * @param <T>
     * @param <U>
     * @param <R>
     * @return
     */
    public static <T,U,R> R biFunction(T t,U u,BiFunction<T,U,R> supplier) {
        return supplier.apply(t,u);
    }
}
