using Xunit;
using WindowsFormsApp1;

namespace TestProject2
{
    public class UnitTest1
    {
        [Fact]
        public void TestValidInput()
        {
            // Arrange

            Form1 form = new Form1();
            form.textBox1.Text = "5";
            form.textBox2.Text = "10";

            // Act
            form.button1_Click(null, null);

            // Assert
            Assert.Equal(1065.0, form.result); // ���������, ��� ��������� ���������� ���������
        }

        [Fact]
        public void TestLargeInput()
        {
            // Arrange
            Form1 form = new Form1();
            form.textBox1.Text = "12345678901234567890";
            form.textBox2.Text = "98765432109876543210";

            // Act
            form.button1_Click(null, null);

            // Assert
            Assert.Equal(1.21822091791791E+39, form.result); // ���������, ��� ��������� ���������� ��������� ��� ������� �����
        }

        [Fact]
        public void TestNegativeInput()
        {
            // Arrange
            Form1 form = new Form1();
            form.textBox1.Text = "-5";
            form.textBox2.Text = "10";

            // Act
            form.button1_Click(null, null);

            // Assert
            Assert.Equal(-663.75, form.result); // ���������, ��� ��������� ���������� ��������� ��� ������������� �����
        }

        [Fact]
        public void TestEmptyInput()
        {
            // Arrange
            Form1 form = new Form1();
            form.textBox1.Text = "";
            form.textBox2.Text = "10";

            // Act
            form.button1_Click(null, null);

            // Assert
            Assert.Equal("���� ������ ���� ���������", form.result); // ���������, ��� ��������� ��������� � ���������� �����
        }
    }
}
