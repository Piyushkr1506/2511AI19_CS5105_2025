package Arrays;
import java.util.Scanner;

/*
Find Second Smallest and Second Largest Element in an array

Problem Statement: Given an array, find the second smallest and second largest element in the array. Print ‘-1’ in the event that either of them doesn’t exist.

Examples

Example 1:
Input: [1,2,4,7,7,5]
Output: Second Smallest : 2
	Second Largest : 5
Explanation: The elements are as follows 1,2,3,5,7,7 and hence second largest of these is 5 and second smallest is 2

Example 2:
Input: [1]
Output: Second Smallest : -1
	Second Largest : -1
Explanation: Since there is only one element in the array, it is the largest and smallest element present in the array. 
There is no second largest or second smallest element present.
 */
public class prob_2 {
    //optimal opproach
    static int secondLargest(int[] arr,int s,int e){
       int largest =arr[s];
       int slargest = -1;
       for(int i=1;i<e;i++){
            if(arr[i]>largest){
                slargest=largest;
                largest=arr[i];
            }
            else if(arr[i]<largest && arr[i]>slargest){
                slargest=arr[i];
            }
       }
       return slargest;
    }
    static int secondSmallest(int[] arr,int s,int e){
       int smallest =arr[s];
       int ssmallest = Integer.MAX_VALUE;
       for(int i=1;i<e;i++){
            if(arr[i]<smallest){
                ssmallest=smallest;
                smallest=arr[i];
            }
            else if(arr[i]!=smallest && arr[i]<ssmallest){
                ssmallest=arr[i];
            }
       }
       return ssmallest;
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }

        int secondLargest = secondLargest(arr,0,N);
        int secondSmallest = secondSmallest(arr,0,N);
        System.out.println("the second largest element :"+secondLargest);
        System.out.println("the second smallest element :"+secondSmallest);
        sc.close();
    }
    
}
