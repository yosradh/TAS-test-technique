
using System;

namespace TAS_Test
{
    partial class Generator
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.generation = new System.Windows.Forms.Button();
            this.pathFichier = new System.Windows.Forms.TextBox();
            this.parcourir = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.generation);
            this.groupBox1.Controls.Add(this.pathFichier);
            this.groupBox1.Controls.Add(this.parcourir);
            this.groupBox1.Location = new System.Drawing.Point(107, 89);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(355, 181);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Importation des données";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // generation
            // 
            this.generation.Location = new System.Drawing.Point(130, 128);
            this.generation.Name = "generation";
            this.generation.Size = new System.Drawing.Size(110, 23);
            this.generation.TabIndex = 2;
            this.generation.Text = "Génération";
            this.generation.UseVisualStyleBackColor = true;
            this.generation.Click += new System.EventHandler(this.Generate_Click);
            // 
            // pathFichier
            // 
            this.pathFichier.Location = new System.Drawing.Point(21, 75);
            this.pathFichier.Name = "pathFichier";
            this.pathFichier.Size = new System.Drawing.Size(328, 20);
            this.pathFichier.TabIndex = 1;
            // 
            // parcourir
            // 
            this.parcourir.Location = new System.Drawing.Point(21, 33);
            this.parcourir.Name = "parcourir";
            this.parcourir.Size = new System.Drawing.Size(92, 23);
            this.parcourir.TabIndex = 0;
            this.parcourir.Text = "Parcourir ...";
            this.parcourir.UseVisualStyleBackColor = true;
            this.parcourir.Click += new System.EventHandler(this.parcourir_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 300);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "File Generator";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button generation;
        private System.Windows.Forms.TextBox pathFichier;
        private System.Windows.Forms.Button parcourir;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

