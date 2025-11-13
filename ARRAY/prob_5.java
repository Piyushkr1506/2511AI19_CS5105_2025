package Arrays;

import java.util.Arrays;
import java.util.Scanner;

/*Left Rotate Array by one.

Problem Statements: Given an integer array nums, rotate the array to the left by one.

Note: There is no need to return anything, just modify the given array.

Example 1;
Input: nums = [1, 2, 3, 4, 5]
Output: [2, 3, 4, 5, 1]
Explanation: Initially, nums = [1, 2, 3, 4, 5]
Rotating once to left -> nums = [2, 3, 4, 5, 1]

Example 2:
Input: nums = [-1, 0, 3, 6]
Output: [0, 3, 6, -1]
Explanation: Initially, nums = [-1, 0, 3, 6]
Rotating once to left -> nums = [0, 3, 6, -1]
 */
public class prob_5 {

    static void reverse(int[] arr,int s,int e){
        while(s<e){
            int temp = arr[s];
            arr[s] =arr[e];
            arr[e] =temp;
            s++;
            e--;
        }
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        //rotation of only one element TC : O(N) SC :O (1) 
        System.out.println("Intitial Array: "+Arrays.toString(arr));
        int last =arr[0];
        for(int i=0;i<N;i++){
            arr[i]=arr[(i+1)%N];
            if(i==N-1){
                arr[N-1]=last;
            }
        }
        
        System.out.println("Array after left rotate by 1:"+Arrays.toString(arr));
        // System.out.println("enter d elements to rotate:");
        // int d = sc.nextInt();
        // reverse(arr, 0, d-1);
        // reverse(arr, d, N-1);
        // reverse(arr, 0, N-1);
        sc.close();
        // System.out.println("Array after left rotate by 1:"+Arrays.toString(arr));
    }
    
}
