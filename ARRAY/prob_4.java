package Arrays;

import java.util.Scanner;

/*Remove duplicates from sorted array.

 Problem Statements: Given an integer array nums sorted in non-decreasing order, remove all duplicates in-place so that each unique element appears only once. 
 Return the number of unique elements in the array.

If the number of unique elements be k, then,

1.Change the array nums such that the first k elements of nums contain the unique values in the order that they were present originally.
The remaining elements, as well as the size of the array does not matter in terms of correctness.


2.An array sorted in non-decreasing order is an array where every element to the right of an element is either equal to or greater in value than that element.
Example 1:
Input: nums = [0, 0, 3, 3, 5, 6]
Output: 4
Explanation: Resulting array = [0, 3, 5, 6, _, _]
There are 4 distinct elements in nums and the elements marked as _ can have any value.

 */
public class prob_4 {
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        //better
        // HashSet<Integer> st =new HashSet<>();
        // for(int i=0;i<N;i++){
        //     st.add(arr[i]);
        // }
        // System.out.println(st.size());
        //optimal approach
        int i=0;
        int j=1;
        while(j<N){
            if(arr[i]==arr[j]){
                j++;
            }
            else{
                arr[i+1]=arr[j];
                i++;
                j++;
            }
        }
        System.out.println("The number of distint elements in array are: "+  (i+1));
        sc.close();

    }
    
}
