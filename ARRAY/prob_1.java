package Arrays;

import java.util.Scanner;

/*Problem Statement: Given an array, we have to find the largest element in the array.

Examples

Example 1:
Input: arr[] = {2,5,1,3,0};
Output: 5
Explanation: 5 is the largest element in the array. 

Example2: 
Input: arr[] = {8,10,5,7,9};
Output: 10
Explanation: 10 is the largest element in the array.
 */
public class prob_1 {
    
    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        int maxElement=arr[0];
        for(int i=1;i<N;i++){
            if(arr[i]>maxElement){
                maxElement=arr[i];
            }
        }
        System.out.println("The Max Element of the given array is: "+maxElement);
        sc.close();
        
    }
}
