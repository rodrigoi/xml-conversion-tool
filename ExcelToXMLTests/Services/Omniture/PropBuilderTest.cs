using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FluentAssertions;
using ExcelToXML.Model;
using ExcelToXML.Services.Omniture;

namespace ExcelToXMLTests.Services.Omniture {
    [TestClass]
    public class PropBuilderTest {

        [TestMethod]
        public void should_have_empty_values_for_empty_reference_object() {
            //Arrange
            var omniture = new ExcelToXML.Model.Omniture();
            dynamic propBuilder = new PropBuilder(omniture);

            //Act
            string pageName = propBuilder.pageName;
            string prop1 = propBuilder.prop1;
            string prop2 = propBuilder.prop2;
            string prop3 = propBuilder.prop3;
            string prop4 = propBuilder.prop4;
            string prop5 = propBuilder.prop5;
            string prop6 = propBuilder.prop6;
            string prop7 = propBuilder.prop7;
            string prop8 = propBuilder.prop8;
            string prop9 = propBuilder.prop9;
            string prop10 = propBuilder.prop10;
            string prop11 = propBuilder.prop11;
            string prop12 = propBuilder.prop12;
            string prop13 = propBuilder.prop13;
            string prop14 = propBuilder.prop14;
            string prop15 = propBuilder.prop15;

            //Assert
            pageName.Should().BeNullOrEmpty();
            prop1.Should().BeNullOrEmpty();
            prop2.Should().BeNullOrEmpty();
            prop3.Should().BeNullOrEmpty();
            prop4.Should().BeNullOrEmpty();
            prop5.Should().BeNullOrEmpty();
            prop6.Should().BeNullOrEmpty();
            prop7.Should().BeNullOrEmpty();
            prop8.Should().BeNullOrEmpty();
            prop9.Should().BeNullOrEmpty();
            prop10.Should().BeNullOrEmpty();
            prop11.Should().BeNullOrEmpty();
            prop12.Should().BeNullOrEmpty();
            prop13.Should().BeNullOrEmpty();
            prop14.Should().BeNullOrEmpty();
            prop15.Should().BeNullOrEmpty();
        }

        [TestMethod]
        public void should_have_all_15_properties() {
            //Arrange
            var omniture = new ExcelToXML.Model.Omniture();
            omniture.Domain = "domain";
            omniture.SiteSection = "site section";
            omniture.SubSection = "sub section";
            omniture.ActiveState = "active state";
            omniture.ObjectValue = "object value";
            omniture.ObjectDescription = "object description";
            omniture.CallToAction = "call to action";
            
            dynamic propBuilder = new PropBuilder(omniture);

            //Act
            string prop1 = propBuilder.prop1;
            string prop2 = propBuilder.prop2;
            string prop3 = propBuilder.prop3;
            string prop4 = propBuilder.prop4;
            string prop5 = propBuilder.prop5;
            string prop6 = propBuilder.prop6;
            string prop7 = propBuilder.prop7;
            string prop8 = propBuilder.prop8;
            string prop9 = propBuilder.prop9;
            string prop10 = propBuilder.prop10;
            string prop11 = propBuilder.prop11;
            string prop12 = propBuilder.prop12;
            string prop13 = propBuilder.prop13;
            string prop14 = propBuilder.prop14;
            string prop15 = propBuilder.prop15;

            //Assert
            prop1.Should().NotBeNullOrEmpty()
                .And.Be("site section");
            prop2.Should().NotBeNullOrEmpty()
                .And.Be("sub section");
            prop3.Should().NotBeNullOrEmpty()
                .And.Be("site section | sub section");
            prop4.Should().NotBeNullOrEmpty()
                .And.Be("active state");
            prop5.Should().NotBeNullOrEmpty()
                .And.Be("sub section | active state");
            prop6.Should().NotBeNullOrEmpty()
                .And.Be("site section | sub section | active state");
            prop7.Should().NotBeNullOrEmpty()
                .And.Be("object value");
            prop8.Should().NotBeNullOrEmpty()
                .And.Be("site section | active state");
            prop9.Should().NotBeNullOrEmpty()
                .And.Be("sub section | active state | object value");
            prop10.Should().NotBeNullOrEmpty()
                .And.Be("site section | sub section | site section | active state");
            prop11.Should().NotBeNullOrEmpty()
                .And.Be("object description");
            prop12.Should().NotBeNullOrEmpty()
                .And.Be("object value | object description");
            prop13.Should().NotBeNullOrEmpty()
                .And.Be("active state | object value | object description");
            prop14.Should().NotBeNullOrEmpty()
                .And.Be("sub section | active state | object value | object description");
            prop15.Should().NotBeNullOrEmpty()
                .And.Be("site section | sub section | active state | object value | object description");
        }

        [TestMethod]
        public void should_build_pageName_for_objectValue_equal_to_ActiveState() {
            //Arrange
            var omniture = new ExcelToXML.Model.Omniture();

            omniture.Domain = "domain";
            omniture.SiteSection = "site section";
            omniture.SubSection = "sub section";
            omniture.ActiveState = "active state";
            omniture.ObjectValue = "active state";
            omniture.ObjectDescription = "object description";
            omniture.CallToAction = "call to action";

            dynamic propBuilder = new PropBuilder(omniture);

            //Act
            string pageName = propBuilder.pageName;

            //Assert
            pageName.Should().NotBeNullOrEmpty()
                .And.Be("domain | site section | sub section | active state | object description");
        }

        [TestMethod]
        public void should_build_pageName_for_objectValue_different_than_ActiveState() {
            //Arrange
            var omniture = new ExcelToXML.Model.Omniture();

            omniture.Domain = "domain";
            omniture.SiteSection = "site section";
            omniture.SubSection = "sub section";
            omniture.ActiveState = "active state";
            omniture.ObjectValue = "object value";
            omniture.ObjectDescription = "object description";
            omniture.CallToAction = "call to action";

            dynamic propBuilder = new PropBuilder(omniture);

            //Act
            string pageName = propBuilder.pageName;

            //Assert
            pageName.Should().NotBeNullOrEmpty()
                .And.Be("domain | site section | sub section | active state | object value | object description");
        }
    }
}