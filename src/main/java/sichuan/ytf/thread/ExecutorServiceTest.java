package sichuan.ytf.thread;

import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class ExecutorServiceTest {

    public static void main(String[] args) throws IOException, InterruptedException {
        // 创建一个固定大小的线程池
        ExecutorService service = Executors.newFixedThreadPool(3);
        for (int i = 0; i < 10; i++) {
            System.out.println("创建线程" + i);
            int j = i;
            Runnable run = new Runnable() {
                @Override
                public void run() {
                    System.out.println("启动线程"+j);
                    try {
                    	if(j==1)
						Thread.sleep(4000);
					} catch (InterruptedException e) {
					}
                }
            };
            // 在未来某个时间执行给定的命令
            service.execute(run);
        }
		do {
			try {
				service.awaitTermination(1000, TimeUnit.MILLISECONDS);
				System.out.println("等待1s");
			} catch (InterruptedException e) {
			}
		} while (!service.isTerminated());
        System.out.println("all thread complete " );
    }
}