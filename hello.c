/*
 * hello.c
 *
 *  Created on: 2018年7月14日
 *      Author: leek
 */
#include<linux/init.h>
#include<linux/module.h>

static int hello_init(void)
{
	printk(KERN_WARNING"HelloWorld\n" );
	return 0;
}
static void hello_exit(void)
{
	printk(KERN_WARNING"GoodBye\n" );
}
module_init(hello_init);
module_exit(hello_exit);
MODULE_LICENSE("GPL" );
MODULE_AUTHOR("liuzhenjun@mail.nankai.edu.cn" );
