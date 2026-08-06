
/*name:farhan shahriar
id:253-35-217
section:46_F1

Task 1: Validate Infix Expression
Write a C program to check the validity of a given infix expression.

Input: A mathematical expression (e.g., (A+B)*(C-D) or A+B*)C().
Output: Display whether the expression is valid or invalid based on balanced parentheses and correct operator placement.
Hint: Use a stack to verify balanced brackets and operator-operand order.

*/

#include <stdio.h>
#include <string.h>
#include <ctype.h>

#define MAX 100


char stack[MAX];
int top = -1;

void push(char ch)
{
    stack[++top] = ch;
}

char pop()
{
    if (top == -1)
        return '\0';
    return stack[top--];
}

int is_Empty()
{
    return top == -1;
}

int is_Operator(char ch)
{
    return (ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '^');
}

int is_Operand(char ch)
{
    return isalnum(ch);   
}

int main()
{
    char exp[MAX];

    printf("Enter an infix expression: ");
    scanf("%s", exp);

    int len = strlen(exp);
    int valid = 1;

    if (is_Operator(exp[0]) || is_Operator(exp[len - 1]))
        valid = 0;

    for (int i = 0; i < len && valid; i++)
    {
        char ch = exp[i];


        if (is_Operand(ch))
        {
            continue;
        }


        else if (ch == '(')
        {
            push(ch);

     
            if (i > 0 && (is_Operand(exp[i - 1]) || exp[i - 1] == ')'))
                valid = 0;
        }

        // Closing bracket
        else if (ch == ')')
        {
            if (is_Empty())
            {
                valid = 0;
                break;
            }

            pop();

            
            if (i > 0 && exp[i - 1] == '(')
                valid = 0;
        }


        else if (is_Operator(ch))
        {
       
            if (i == 0 || i == len - 1)
                valid = 0;

            else if (is_Operator(exp[i - 1]) || exp[i - 1] == '(')
                valid = 0;

            else if (is_Operator(exp[i + 1]) || exp[i + 1] == ')')
                valid = 0;
        }


        else
        {
            valid = 0;
        }
    }


    if (!is_Empty())
        valid = 0;

    if (valid)
        printf("\nValid Infix Expression\n");
    else
        printf("\nInvalid Infix Expression\n");

    return 0;
}