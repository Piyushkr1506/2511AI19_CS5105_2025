package Arrays;

import java.util.Arrays;
import java.util.Scanner;

/*Move Zeros to End.

Problem Statements: Given an integer array nums, move all the 0's to the end of the array. 
The relative order of the other elements must remain the same. This must be done in place, without making a copy of the array.

Example 1:
Input: nums = [0, 1, 4, 0, 5, 2]

Output: [1, 4, 5, 2, 0, 0]
 */
public class prob_6 {
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        System.out.println("Initial Array:");
        System.out.println(Arrays.toString(arr));
        int j=0;
        for(int i=0;i<N;i++){
            if(arr[i]!=0){
                arr[j]=arr[i];
                j++;
            }
        }
        while(j<N){
            arr[j]=0;
            j++;
        }
        System.out.println("Final Array after moving Zeroes to end:");
       System.out.println(Arrays.toString(arr));
       sc.close();
    }
    
}
